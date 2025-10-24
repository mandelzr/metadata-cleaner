import os
import re
import shutil
import struct
import tempfile
from pathlib import Path
import sys
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import json
import subprocess
import hashlib

try:
    import pikepdf  # type: ignore
    _HAS_PIKEPDF = True
except Exception:
    pikepdf = None  # type: ignore
    _HAS_PIKEPDF = False

# Optional Windows OLE Structured Storage (for legacy Office)
try:
    import pythoncom  # type: ignore
    import win32com.storagecon as storagecon  # type: ignore
    _HAS_PYWIN32 = True
except Exception:
    pythoncom = None  # type: ignore
    storagecon = None  # type: ignore
    _HAS_PYWIN32 = False


def _magic(path, n=16):
    try:
        with open(path, 'rb') as f:
            return f.read(n)
    except Exception:
        return b''


def _file_type(path):
    p = Path(path)
    ext = p.suffix.lower()
    head = _magic(path, 64)
    if head.startswith(b"\xFF\xD8\xFF"):
        return 'jpeg'
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return 'png'
    if head.startswith(b"GIF87a") or head.startswith(b"GIF89a"):
        return 'gif'
    if head.startswith(b"%PDF-"):
        return 'pdf'
    if head.startswith(b"{\\rtf"):
        return 'rtf'
    # OOXML stored in ZIP but with legacy extensions
    if head.startswith(b'PK') and ext in ('.doc', '.xls', '.ppt'):
        ooxml = _detect_ooxml_from_zip(path)
        if ooxml:
            return ooxml
    if ext == '.doc':
        return 'doc'
    if ext == '.xls':
        return 'xls'
    if ext == '.ppt':
        return 'ppt'
    if ext in ('.docx', '.xlsx', '.pptx'):
        return ext[1:]
    # Word 2003 XML sometimes mislabeled as .doc (plain XML)
    try:
        head_strip = head.lstrip()
        if head_strip.startswith(b'<?xml') or head_strip.startswith(b'<'):
            if _is_word2003xml(path):
                return 'word2003xml'
    except Exception:
        pass
    return 'other'


def _detect_ooxml_from_zip(path):
    try:
        with ZipFile(path, 'r') as z:
            names = set(z.namelist())
            if 'word/document.xml' in names:
                return 'docx'
            if 'xl/workbook.xml' in names:
                return 'xlsx'
            if 'ppt/presentation.xml' in names:
                return 'pptx'
    except Exception:
        return None
    return None


def _is_word2003xml(path: str) -> bool:
    try:
        with open(path, 'rb') as f:
            data = f.read(8192)
        text = data.decode('utf-8', errors='ignore')
        return ('w:wordDocument' in text) or ('wordml' in text and 'http://schemas.microsoft.com/office/word' in text)
    except Exception:
        return False


def detect_file_metadata(path):
    t = _file_type(path)
    result = {"type": t, "can_clean": False, "summary": []}
    if t == 'word2003xml':
        props = _detect_word2003xml_props(path)
        result["summary"] = props
        result["can_clean"] = bool(props)
        return result
    if t == 'rtf':
        tags = _detect_rtf_info(path)
        result["summary"] = tags
        result["can_clean"] = bool(tags)
        return result
    # Legacy Office binary formats: prefer native OLE property detection/cleaning
    if t in ('doc', 'xls', 'ppt'):
        # Probe OLE property streams
        has_sum, has_docsum = _ole_has_props(path)
        if has_sum:
            result["summary"].append('SummaryInfo')
        if has_docsum:
            result["summary"].append('DocSummaryInfo')
        # Also consult ExifTool to surface fields even if OLE detection fails
        try:
            exiftool = _find_exiftool()
            if exiftool:
                et = _exiftool_detect_summary(exiftool, path)
                if et:
                    for tname in et:
                        if tname not in result["summary"]:
                            result["summary"].append(tname)
        except Exception:
            pass
        # We can attempt cleaning if pywin32 is available; success depends on streams existing
        result["can_clean"] = bool(_HAS_PYWIN32)
        if not (has_sum or has_docsum) and result["summary"] and 'Metadata detected' not in result["summary"]:
            # If ExifTool saw fields but we didn't find the streams, show a generic note in summary
            pass
        if not _HAS_PYWIN32:
            result["note"] = "Legacy Office cleaning requires Windows pywin32 (bundled in EXE)"
        return result
    # Prefer ExifTool when available (broadest coverage)
    exiftool = _find_exiftool()
    if exiftool:
        et_summary = _exiftool_detect_summary(exiftool, path)
        if et_summary is not None:
            result["summary"] = et_summary
            result["can_clean"] = len(et_summary) > 0
            # For unknown types, display the real extension instead of 'exiftool'
            if t == 'other':
                ext = Path(path).suffix.lower().lstrip('.')
                result["type"] = ext if ext else 'other'
            return result
    if t == 'jpeg':
        exif, xmp, iptc = _detect_jpeg(path)
        result["summary"] = [s for s, present in (("EXIF", exif), ("XMP", xmp), ("IPTC", iptc)) if present]
        result["can_clean"] = bool(result["summary"])  # only if something present
        return result
    if t == 'png':
        textc, timec = _detect_png(path)
        if textc:
            result["summary"].append(f"Text chunks:{textc}")
        if timec:
            result["summary"].append("tIME")
        result["can_clean"] = bool(textc or timec)
        return result
    if t == 'gif':
        comments = _detect_gif_comments(path)
        if comments:
            result["summary"].append(f"Comments:{comments}")
        result["can_clean"] = comments > 0
        return result
    if t in ('docx', 'xlsx', 'pptx'):
        info = _detect_office_props_details(path)
        result["summary"] = info["summary"]
        result["can_clean"] = info["present"]
        return result
    if t == 'pdf':
        if _HAS_PIKEPDF:
            info = _detect_pdf_metadata_pike(path)
            result["summary"] = info["summary"]
            result["can_clean"] = info["present"]
            if not result["summary"] and info["present"]:
                result["summary"] = ["Info/XMP present"]
            return result
        else:
            has_meta = _detect_pdf_metadata_quick(path)
            result["summary"] = ["Metadata detected"] if has_meta else []
            result["can_clean"] = False
            result["note"] = "PDF cleaning requires pikepdf (not bundled)"
            return result
    if t == 'ppt':
        result["note"] = "Legacy PPT not supported; PPTX is supported"
        result["can_clean"] = False
        return result
    result["note"] = "Unsupported format"
    return result


def clean_file_metadata(path, backup=True):
    t = _file_type(path)
    if t == 'word2003xml':
        changed = _clean_word2003xml(path, backup)
        return changed, "Removed Office 2003 XML properties" if changed else ""
    if t == 'rtf':
        changed = _clean_rtf(path, backup)
        return changed, "Removed RTF \\info" if changed else ""
    # Legacy Office binary formats: remove OLE SummaryInformation / DocumentSummaryInformation
    if t in ('doc', 'xls', 'ppt'):
        if not _HAS_PYWIN32:
            return False, "Legacy Office cleaning unavailable (missing pywin32)"
        changed, reason = _clean_ole_props(path, backup)
        return changed, ("Removed OLE property sets" if changed else reason)
    # Prefer native Office cleaner for OOXML (docx/xlsx/pptx) because ExifTool
    # does not support writing these containers.
    if t in ('docx', 'xlsx', 'pptx'):
        removed = _clean_office_props(path, backup)
        return removed > 0, f"Removed docProps ({removed})" if removed else ""
    # Prefer ExifTool for cleaning when available
    exiftool = _find_exiftool()
    if exiftool:
        bak_path = None
        if backup:
            try:
                bak_path = _make_backup_copy(path)
            except Exception:
                bak_path = None
        changed, reason = _exiftool_clean(exiftool, path, t)
        if not changed and bak_path:
            try:
                os.remove(bak_path)
            except Exception:
                pass
        if changed:
            return True, "ExifTool: removed metadata"
        else:
            return False, (reason or "")
    if t == 'jpeg':
        changed = _clean_jpeg(path, backup)
        return changed, "Removed EXIF/XMP/IPTC" if changed else ""
    if t == 'png':
        changed = _clean_png(path, backup)
        return changed, "Removed text/time chunks" if changed else ""
    if t == 'gif':
        changed = _clean_gif(path, backup)
        return changed, "Removed comments" if changed else ""
    # OOXML handled above
    if t == 'pdf' and _HAS_PIKEPDF:
        changed = _clean_pdf(path, backup)
        return changed, "Removed Info/XMP" if changed else ""
    return False, ""


# ---------------- JPEG ----------------


def _detect_jpeg(path):
    exif = xmp = iptc = False
    with open(path, 'rb') as f:
        data = f.read(2)
        if data != b"\xFF\xD8":
            return False, False, False
        while True:
            marker = f.read(2)
            if len(marker) < 2:
                break
            if marker[0] != 0xFF:
                break
            if marker[1] == 0xDA:  # SOS
                break
            if marker[1] in (0xD8, 0xD9):
                continue
            length_bytes = f.read(2)
            if len(length_bytes) < 2:
                break
            length = struct.unpack('>H', length_bytes)[0]
            seg_data = f.read(length - 2)
            if marker[1] == 0xE1:  # APP1
                if seg_data.startswith(b'Exif\x00\x00'):
                    exif = True
                if b'http://ns.adobe.com/xap/1.0/' in seg_data:
                    xmp = True
            if marker[1] == 0xED:  # APP13 IPTC
                if seg_data.startswith(b'Photoshop 3.0'):
                    iptc = True
    return exif, xmp, iptc


def _clean_jpeg(path, backup=True):
    changed = False
    with open(path, 'rb') as f:
        data = f.read(2)
        if data != b"\xFF\xD8":
            return False
        tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=".jpg", dir=str(Path(path).parent))
        os.close(tmp_fd)
        with open(tmp_path, 'wb') as out:
            out.write(b"\xFF\xD8")
            while True:
                marker = f.read(2)
                if len(marker) < 2:
                    break
                if marker[0] != 0xFF:
                    break
                if marker[1] == 0xDA:  # SOS
                    # copy rest including SOS marker and compressed data
                    out.write(marker)
                    rest = f.read()
                    out.write(rest)
                    break
                if marker[1] in (0xD8, 0xD9):
                    out.write(marker)
                    continue
                length_bytes = f.read(2)
                if len(length_bytes) < 2:
                    break
                length = struct.unpack('>H', length_bytes)[0]
                seg_data = f.read(length - 2)
                drop = False
                if marker[1] == 0xE1 and (seg_data.startswith(b'Exif\x00\x00') or b'http://ns.adobe.com/xap/1.0/' in seg_data):
                    drop = True
                if marker[1] == 0xED and seg_data.startswith(b'Photoshop 3.0'):
                    drop = True
                if drop:
                    changed = True
                else:
                    out.write(marker)
                    out.write(length_bytes)
                    out.write(seg_data)
    if changed:
        _replace_file(path, tmp_path, backup)
    else:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    return changed


# ---------------- PNG ----------------


def _detect_png(path):
    text_chunks = 0
    time_chunks = 0
    with open(path, 'rb') as f:
        sig = f.read(8)
        if sig != b"\x89PNG\r\n\x1a\n":
            return 0, 0
        while True:
            len_bytes = f.read(4)
            if len(len_bytes) < 4:
                break
            length = struct.unpack('>I', len_bytes)[0]
            type_bytes = f.read(4)
            if len(type_bytes) < 4:
                break
            chunk_type = type_bytes
            f.seek(length + 4, os.SEEK_CUR)  # skip data + crc
            if chunk_type in (b'tEXt', b'iTXt', b'zTXt'):
                text_chunks += 1
            if chunk_type == b'tIME':
                time_chunks += 1
    return text_chunks, time_chunks


def _clean_png(path, backup=True):
    changed = False
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=".png", dir=str(Path(path).parent))
    os.close(tmp_fd)
    with open(path, 'rb') as f, open(tmp_path, 'wb') as out:
        sig = f.read(8)
        if sig != b"\x89PNG\r\n\x1a\n":
            os.remove(tmp_path)
            return False
        out.write(sig)
        while True:
            len_bytes = f.read(4)
            if len(len_bytes) < 4:
                break
            length = struct.unpack('>I', len_bytes)[0]
            ctype = f.read(4)
            data = f.read(length)
            crc = f.read(4)
            if ctype in (b'tEXt', b'iTXt', b'zTXt', b'tIME'):
                changed = True
                continue  # drop
            out.write(len_bytes)
            out.write(ctype)
            out.write(data)
            out.write(crc)
    if changed:
        _replace_file(path, tmp_path, backup)
    else:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    return changed


# ---------------- GIF ----------------


def _detect_gif_comments(path):
    comments = 0
    with open(path, 'rb') as f:
        hdr = f.read(6)
        if hdr not in (b'GIF87a', b'GIF89a'):
            return 0
        # Logical Screen Descriptor
        lsd = f.read(7)
        if len(lsd) < 7:
            return 0
        packed = lsd[4]
        gct_flag = (packed & 0x80) != 0
        gct_size = 3 * (2 ** ((packed & 0x07) + 1)) if gct_flag else 0
        f.seek(gct_size, os.SEEK_CUR)
        # Blocks
        while True:
            introducer = f.read(1)
            if not introducer:
                break
            b = introducer[0]
            if b == 0x3B:  # trailer
                break
            if b == 0x2C:  # image descriptor
                f.seek(9, os.SEEK_CUR)
                packed = f.read(1)
                if not packed:
                    break
                lct_flag = (packed[0] & 0x80) != 0
                lct_size = 3 * (2 ** ((packed[0] & 0x07) + 1)) if lct_flag else 0
                f.seek(lct_size, os.SEEK_CUR)  # local color table
                # LZW min code size
                f.seek(1, os.SEEK_CUR)
                # image data sub-blocks
                _skip_sub_blocks(f)
            elif b == 0x21:  # extension
                label = f.read(1)
                if not label:
                    break
                if label[0] == 0xFE:  # comment extension
                    comments += 1
                    _skip_sub_blocks(f)
                else:
                    _skip_sub_blocks(f)
            else:
                break
    return comments


def _clean_gif(path, backup=True):
    changed = False
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=".gif", dir=str(Path(path).parent))
    os.close(tmp_fd)
    with open(path, 'rb') as f, open(tmp_path, 'wb') as out:
        hdr = f.read(6)
        if hdr not in (b'GIF87a', b'GIF89a'):
            os.remove(tmp_path)
            return False
        out.write(hdr)
        lsd = f.read(7)
        if len(lsd) < 7:
            os.remove(tmp_path)
            return False
        out.write(lsd)
        packed = lsd[4]
        gct_flag = (packed & 0x80) != 0
        gct_size = 3 * (2 ** ((packed & 0x07) + 1)) if gct_flag else 0
        if gct_size:
            out.write(f.read(gct_size))
        while True:
            introducer = f.read(1)
            if not introducer:
                break
            b = introducer[0]
            out_pos_written = False
            if b == 0x3B:  # trailer
                out.write(introducer)
                break
            if b == 0x2C:  # image descriptor
                out.write(introducer)
                out.write(f.read(9))
                packed2 = f.read(1)
                out.write(packed2)
                lct_flag = (packed2[0] & 0x80) != 0
                lct_size = 3 * (2 ** ((packed2[0] & 0x07) + 1)) if lct_flag else 0
                if lct_size:
                    out.write(f.read(lct_size))
                # LZW min code size
                out.write(f.read(1))
                # image data sub-blocks copy
                _copy_sub_blocks(f, out)
                continue
            if b == 0x21:  # extension
                label = f.read(1)
                if not label:
                    break
                if label[0] == 0xFE:  # comment extension - drop
                    changed = True
                    _skip_sub_blocks(f)
                    continue
                else:
                    out.write(introducer)
                    out.write(label)
                    _copy_sub_blocks(f, out)
                    continue
            # unknown -> break to avoid corruption
            break
    if changed:
        _replace_file(path, tmp_path, backup)
    else:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
    return changed


def _skip_sub_blocks(f):
    while True:
        szb = f.read(1)
        if not szb:
            break
        sz = szb[0]
        if sz == 0:
            break
        f.seek(sz, os.SEEK_CUR)


def _copy_sub_blocks(f, out):
    while True:
        szb = f.read(1)
        if not szb:
            break
        out.write(szb)
        sz = szb[0]
        if sz == 0:
            break
        out.write(f.read(sz))


# ---------------- Office (docx/xlsx/pptx) ----------------


def _detect_office_props(path):
    prop_paths = {
        'core.xml': 'docProps/core.xml',
        'app.xml': 'docProps/app.xml',
        'custom.xml': 'docProps/custom.xml',
    }
    found = {k: False for k in prop_paths}
    try:
        with ZipFile(path, 'r') as z:
            names = set(z.namelist())
            for k, v in prop_paths.items():
                if v in names:
                    found[k] = True
        return found
    except Exception:
        return found


def _detect_office_props_details(path):
    result = {"present": False, "summary": []}
    try:
        with ZipFile(path, 'r') as z:
            names = set(z.namelist())
            # core.xml
            if 'docProps/core.xml' in names:
                try:
                    core = ET.fromstring(z.read('docProps/core.xml'))
                    ns = {
                        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                        'dc': 'http://purl.org/dc/elements/1.1/',
                        'dcterms': 'http://purl.org/dc/terms/'
                    }
                    if core.find('dc:creator', ns) is not None:
                        result["summary"].append('Author')
                    if core.find('cp:lastModifiedBy', ns) is not None:
                        result["summary"].append('LastModifiedBy')
                    if core.find('dcterms:created', ns) is not None:
                        result["summary"].append('Created')
                    if core.find('dcterms:modified', ns) is not None:
                        result["summary"].append('Modified')
                    if core.find('dc:title', ns) is not None:
                        result["summary"].append('Title')
                    if core.find('dc:subject', ns) is not None:
                        result["summary"].append('Subject')
                    if core.find('cp:keywords', ns) is not None:
                        result["summary"].append('Keywords')
                    if core.find('cp:category', ns) is not None:
                        result["summary"].append('Category')
                except Exception:
                    result["summary"].append('CoreProps')
                result["present"] = True
            # app.xml
            if 'docProps/app.xml' in names:
                try:
                    app = ET.fromstring(z.read('docProps/app.xml'))
                    ns2 = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
                    if app.find('ep:Company', ns2) is not None:
                        result["summary"].append('Company')
                    if app.find('ep:Manager', ns2) is not None:
                        result["summary"].append('Manager')
                    if app.find('ep:Application', ns2) is not None:
                        result["summary"].append('Application')
                except Exception:
                    result["summary"].append('AppProps')
                result["present"] = True
            # custom.xml
            if 'docProps/custom.xml' in names:
                try:
                    custom = ET.fromstring(z.read('docProps/custom.xml'))
                    props = custom.findall('.//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property')
                    if props:
                        result["summary"].append(f'CustomProps:{len(props)}')
                except Exception:
                    result["summary"].append('CustomProps')
                result["present"] = True
            # thumbnails
            thumbs = [n for n in names if n.startswith('docProps/thumbnail.')]
            if thumbs:
                result["summary"].append('Thumbnail')
                result["present"] = True
    except Exception:
        pass
    return result


def _clean_office_props(path, backup=True):
    remove = {'docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'}
    removed = 0
    dirp = str(Path(path).parent)
    tmp_dir = tempfile.mkdtemp(prefix="_mc_zip_", dir=dirp)
    tmp_path = os.path.join(tmp_dir, Path(path).name)
    try:
        with ZipFile(path, 'r') as zin, ZipFile(tmp_path, 'w', ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename in remove or item.filename.startswith('docProps/thumbnail.'):
                    removed += 1
                    continue
                data = zin.read(item.filename)
                # Clean references in content types
                if item.filename == '[Content_Types].xml':
                    try:
                        xml = ET.fromstring(data)
                        ns_ct = '{http://schemas.openxmlformats.org/package/2006/content-types}'
                        for ov in list(xml.findall(f'{ns_ct}Override')):
                            part = ov.attrib.get('PartName', '')
                            if part.startswith('/docProps/'):
                                xml.remove(ov)
                        data = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass
                # Clean references in root relationships
                elif item.filename == '_rels/.rels':
                    try:
                        xml = ET.fromstring(data)
                        ns_rel = '{http://schemas.openxmlformats.org/package/2006/relationships}'
                        for rel in list(xml.findall(f'{ns_rel}Relationship')):
                            target = rel.attrib.get('Target', '')
                            if target.startswith('docProps/'):
                                xml.remove(rel)
                        data = ET.tostring(xml, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass
                zout.writestr(item, data)
        if removed:
            _replace_file(path, tmp_path, backup)
        else:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    finally:
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
    return removed


# ---------------- PDF (detect only) ----------------


def _detect_pdf_metadata_quick(path):
    try:
        with open(path, 'rb') as f:
            data = f.read(65536)
            if b'/Metadata' in data or b'xpacket' in data or b'/Info' in data:
                return True
    except Exception:
        pass
    return False


def _detect_pdf_metadata_pike(path):
    summary = []
    present = False
    try:
        assert _HAS_PIKEPDF and pikepdf is not None
        with pikepdf.open(path) as pdf:  # type: ignore
            try:
                for k, v in pdf.docinfo.items():
                    key = str(k).lstrip('/')
                    if key:
                        present = True
                        if key in ("Title", "Author", "Subject", "Keywords", "Creator", "Producer", "CreationDate", "ModDate"):
                            if key not in summary:
                                summary.append(key)
                        else:
                            if "CustomInfo" not in summary:
                                summary.append("CustomInfo")
            except Exception:
                pass
            md = getattr(pdf.root, 'Metadata', None)
            if md is not None:
                present = True
                try:
                    data = bytes(md.read_bytes())
                except Exception:
                    data = b''
                if data:
                    xmp_text = data.decode('utf-8', errors='ignore')
                    tags = [
                        ("dc:title", "Title"),
                        ("dc:creator", "Author"),
                        ("xmp:CreatorTool", "CreatorTool"),
                        ("pdf:Producer", "Producer"),
                        ("xmp:CreateDate", "CreateDate"),
                        ("xmp:ModifyDate", "ModifyDate"),
                        ("xmpMM:DocumentID", "DocumentID"),
                    ]
                    for needle, label in tags:
                        if needle in xmp_text and label not in summary:
                            summary.append(label)
                elif "XMP" not in summary:
                    summary.append("XMP")
        return {"present": present, "summary": summary}
    except Exception:
        return {"present": _detect_pdf_metadata_quick(path), "summary": ["Metadata detected"]}


def _clean_pdf(path, backup=True):
    changed = False
    dirp = str(Path(path).parent)
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=".pdf", dir=dirp)
    os.close(tmp_fd)
    try:
        assert _HAS_PIKEPDF and pikepdf is not None
        with pikepdf.open(path, allow_overwriting_input=False) as pdf:  # type: ignore
            if pdf.docinfo:
                pdf.docinfo.clear()
                changed = True
            if getattr(pdf.root, 'Metadata', None) is not None:
                try:
                    del pdf.root.Metadata
                except Exception:
                    pdf.root.Metadata = None
                changed = True
            pdf.save(tmp_path)
        if changed:
            _replace_file(path, tmp_path, backup)
        else:
            os.remove(tmp_path)
    except Exception:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise
    return changed


# ---------------- File replace helpers ----------------


def _replace_file(orig, new_tmp, backup):
    orig = Path(orig)
    new_tmp = Path(new_tmp)
    if backup:
        bak = orig.with_suffix(orig.suffix + '.bak')
        idx = 1
        while bak.exists():
            bak = orig.with_suffix(orig.suffix + f'.bak.{idx}')
            idx += 1
        try:
            shutil.copy2(str(orig), str(bak))
        except Exception:
            pass
    os.replace(str(new_tmp), str(orig))


def _make_backup_copy(orig_path: str) -> str:
    orig = Path(orig_path)
    bak = orig.with_suffix(orig.suffix + '.bak')
    idx = 1
    while bak.exists():
        bak = orig.with_suffix(orig.suffix + f'.bak.{idx}')
        idx += 1
    shutil.copy2(str(orig), str(bak))
    return str(bak)


# ---------------- ExifTool backend ----------------


def _find_exiftool():
    candidates = [
        'exiftool',
        str(Path.home() / 'AppData/Local/Programs/ExifTool/ExifTool.exe'),
        str(Path('C:/Users') / os.getenv('USERNAME', '') / 'AppData/Local/Programs/ExifTool/ExifTool.exe'),
    ]
    # Check for PyInstaller-bundled binary
    try:
        meipass = getattr(sys, '_MEIPASS', None)
    except Exception:
        meipass = None
    if meipass:
        bundled = Path(meipass) / 'exiftool' / 'ExifTool.exe'
        candidates.insert(0, str(bundled))
        bundled2 = Path(meipass) / 'ExifTool.exe'
        candidates.insert(0, str(bundled2))
    # Check relative to script (for dev runs)
    here = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    rel = here / 'exiftool' / 'ExifTool.exe'
    candidates.insert(0, str(rel))
    for c in candidates:
        try:
            proc = subprocess.run([c, '-ver'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=3, **_subprocess_hide_console())
            if proc.returncode == 0 and proc.stdout.strip():
                return c
        except Exception:
            continue
    return None


def exiftool_sensitive_labels(path: str):
    """Return a list of commonly sensitive tags reported by ExifTool.

    Uses the same filtering as our detection helper. Returns [] if ExifTool
    is unavailable or returns no data.
    """
    exiftool = _find_exiftool()
    if not exiftool:
        return []
    labels = _exiftool_detect_summary(exiftool, path)
    return labels or []


def ole_props_state(path: str):
    """Return (has_SummaryInfo, has_DocSummaryInfo) for legacy OLE files.

    If pywin32 is unavailable or file is not OLE, returns (False, False).
    """
    if not _HAS_PYWIN32:
        return (False, False)
    try:
        return _ole_has_props(path)
    except Exception:
        return (False, False)


def _exiftool_detect_summary(exiftool, path):
    try:
        proc = subprocess.run([exiftool, '-j', '-a', '-G1', '-s', path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=20, **_subprocess_hide_console())
        if proc.returncode != 0:
            return None
        data = json.loads(proc.stdout)
        if not data:
            return []
        tags = data[0]
        # Only show commonly sensitive tags for readability
        sensitive_keys = {
            'Author', 'Creator', 'Producer', 'Title', 'Subject', 'Keywords', 'CreatorTool', 'CreateDate', 'ModifyDate',
            'LastModifiedBy', 'Company', 'Manager', 'Category', 'DocSecurity', 'Application', 'OwnerName', 'Artist',
            'Copyright', 'XPAuthor', 'XPComment', 'XPKeywords', 'Make', 'Model', 'GPSLatitude', 'GPSLongitude',
        }
        present = []
        for k, v in tags.items():
            # k may be like 'XMP:CreatorTool' or 'PDF:Producer'
            label = k.split(':', 1)[-1]
            if label in sensitive_keys and v not in (None, '', 0, '0'):
                if label not in present:
                    present.append(label)
        return present
    except Exception:
        return None


def _exiftool_clean(exiftool, path, filetype=None):
    # Use -all= to remove all writable metadata; -overwrite_original to keep filename
    # For safety, we already create backups in our own routines; here we rely on overwrite
    try:
        cmd = [exiftool, '-overwrite_original']
        if filetype == 'doc':
            # Be explicit for legacy OLE Word docs
            cmd += ['-SummaryInfo:All=', '-DocSummaryInfo:All=']
        else:
            cmd += ['-all=']
        cmd.append(path)
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120, **_subprocess_hide_console())
        out = (proc.stdout or '') + '\n' + (proc.stderr or '')
        out_l = out.lower()
        if proc.returncode == 0:
            if 'updated' in out_l:
                return True, ''
            if 'unchanged' in out_l or 'nothing to do' in out_l:
                return False, 'No writable metadata for this file type'
            # Sometimes ExifTool reports success without keywords; fall back to not-changed with note
            return False, 'No changes reported by ExifTool'
        else:
            return False, f'ExifTool error: {out.strip()}'
    except Exception as e:
        return False, f'ExifTool exception: {e}'


# ---------------- Legacy Office (OLE Structured Storage) ----------------


def _ole_has_props(path: str) -> tuple[bool, bool]:
    if not _HAS_PYWIN32:
        return False, False
    try:
        # Open storage read-only and probe streams
        stg = pythoncom.StgOpenStorage(path, None,
                                       storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE,
                                       None, 0)  # type: ignore
        has_sum = _ole_stream_exists(stg, u"\x05SummaryInformation")
        has_docsum = _ole_stream_exists(stg, u"\x05DocumentSummaryInformation")
        stg = None
        return bool(has_sum), bool(has_docsum)
    except Exception:
        return False, False


def _ole_stream_exists(istg, name: str) -> bool:
    try:
        istg.OpenStream(name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE, 0)  # type: ignore
        return True
    except Exception:
        return False


def _clean_ole_props(path: str, backup: bool) -> tuple[bool, str]:
    try:
        # Make backup first if requested
        if backup:
            try:
                _make_backup_copy(path)
            except Exception:
                pass
        # Open storage read-write
        stg = pythoncom.StgOpenStorage(path, None,
                                       storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE,
                                       None, 0)  # type: ignore
        changed = False
        for name in (u"\x05SummaryInformation", u"\x05DocumentSummaryInformation"):
            try:
                stg.DestroyElement(name)
                changed = True
            except Exception:
                pass
        # Commit changes
        try:
            stg.Commit(0)  # type: ignore
        except Exception:
            pass
        stg = None
        if changed:
            return True, ''
        else:
            return False, 'No OLE property sets present'
    except Exception as e:
        return False, f'OLE clean error: {e}'


# ---------------- RTF ----------------


def _read_text_latin1(path: str) -> str:
    with open(path, 'r', encoding='latin-1', errors='ignore') as f:
        return f.read()


def _detect_rtf_info(path: str):
    try:
        s = _read_text_latin1(path)
    except Exception:
        return []
    tags = []
    for block in _rtf_info_blocks(s):
        for key, label in (
            ('\\author', 'Author'),
            ('\\company', 'Company'),
            ('\\title', 'Title'),
            ('\\subject', 'Subject'),
            ('\\keywords', 'Keywords'),
            ('\\operator', 'Operator'),
            ('\\category', 'Category'),
            ('\\doccomm', 'Comment'),
            ('\\creatim', 'CreateTime'),
            ('\\revtim', 'ModTime'),
        ):
            if key in block and label not in tags:
                tags.append(label)
    return tags


def _clean_rtf(path: str, backup: bool) -> bool:
    try:
        s = _read_text_latin1(path)
    except Exception:
        return False
    stripped, changed = _rtf_strip_info(s)
    if not changed:
        return False
    dirp = str(Path(path).parent)
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=".rtf", dir=dirp)
    os.close(tmp_fd)
    with open(tmp_path, 'w', encoding='latin-1', errors='ignore', newline='\r\n') as out:
        out.write(stripped)
    _replace_file(path, tmp_path, backup)
    return True


def _rtf_info_blocks(s: str):
    blocks = []
    i = 0
    n = len(s)
    while i < n:
        if s[i] == '{' and s.startswith('\\info', i + 1):
            # enter group, track braces
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                c = s[j]
                if c == '\\':
                    j += 2
                    continue
                if c == '{':
                    depth += 1
                elif c == '}':
                    depth -= 1
                j += 1
            blocks.append(s[i:j])
            i = j
            continue
        i += 1
    return blocks


def _rtf_strip_info(s: str) -> tuple[str, bool]:
    out = []
    i = 0
    n = len(s)
    changed = False
    while i < n:
        if s[i] == '{' and s.startswith('\\info', i + 1):
            # skip this group
            depth = 1
            j = i + 1
            while j < n and depth > 0:
                c = s[j]
                if c == '\\':
                    j += 2
                    continue
                if c == '{':
                    depth += 1
                elif c == '}':
                    depth -= 1
                j += 1
            i = j
            changed = True
            continue
        # normal copy, handle escapes
        if s[i] == '\\' and i + 1 < n:
            out.append(s[i])
            out.append(s[i + 1])
            i += 2
        else:
            out.append(s[i])
            i += 1
    return (''.join(out), changed)


def _hash_rtf_content(path: str) -> str:
    try:
        s = _read_text_latin1(path)
        stripped, _ = _rtf_strip_info(s)
        h = hashlib.sha256()
        h.update(stripped.encode('latin-1', errors='ignore'))
        return h.hexdigest()
    except Exception:
        return _hash_file(path)


# ---------------- Word 2003 XML ----------------


def _detect_word2003xml_props(path: str):
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        # Namespaces commonly used in Word 2003 XML
        ns = {
            'o': 'urn:schemas-microsoft-com:office:office',
            'w': 'http://schemas.microsoft.com/office/word/2003/wordml',
            'v': 'urn:schemas-microsoft-com:vml',
            'dc': 'http://purl.org/dc/elements/1.1/'
        }
        tags = []
        dp = root.find('.//o:DocumentProperties', ns)
        if dp is not None:
            for child in list(dp):
                label = child.tag.split('}', 1)[-1]
                if label not in tags:
                    tags.append(label)
        cdp = root.find('.//o:CustomDocumentProperties', ns)
        if cdp is not None:
            tags.append('CustomDocumentProperties')
        return tags
    except Exception:
        return []


def _clean_word2003xml(path: str, backup: bool) -> bool:
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        ns = {
            'o': 'urn:schemas-microsoft-com:office:office',
            'w': 'http://schemas.microsoft.com/office/word/2003/wordml',
        }
        changed = False
        dp = root.find('.//o:DocumentProperties', ns)
        if dp is not None:
            parent = root.find('.//o:DocumentProperties/..', ns)
            try:
                # If we can’t get parent via xpath consistently, remove via manual search
                for elem in list(root.iter()):
                    for child in list(elem):
                        if child is dp:
                            elem.remove(child)
                            changed = True
                            raise StopIteration
            except StopIteration:
                pass
        cdp = root.find('.//o:CustomDocumentProperties', ns)
        if cdp is not None:
            try:
                for elem in list(root.iter()):
                    for child in list(elem):
                        if child is cdp:
                            elem.remove(child)
                            changed = True
                            raise StopIteration
            except StopIteration:
                pass
        if not changed:
            return False
        # Write back to temp then replace
        dirp = str(Path(path).parent)
        tmp_fd, tmp_path = tempfile.mkstemp(prefix="_mc_", suffix=Path(path).suffix, dir=dirp)
        os.close(tmp_fd)
        tree.write(tmp_path, encoding='utf-8', xml_declaration=True)
        _replace_file(path, tmp_path, backup)
        return True
    except Exception:
        return False


def _subprocess_hide_console():
    """Return kwargs to hide console windows for subprocess on Windows."""
    if os.name == 'nt':
        kwargs: dict = {}
        try:
            si = subprocess.STARTUPINFO()
            si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            kwargs['startupinfo'] = si
        except Exception:
            pass
        try:
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        except Exception:
            pass
        return kwargs
    return {}


# ---------------- Content hash (verification) ----------------


def compute_content_hash(path: str):
    """Compute a content-only SHA-256 hash for supported formats.

    Returns a tuple (hex_digest, description) or (None, reason) if not available.
    """
    ft = _file_type(path)
    try:
        if ft == 'jpeg':
            return _hash_jpeg_scan(path), 'JPEG scan data'
        if ft == 'png':
            return _hash_png_idat(path), 'PNG IDAT'
        if ft == 'gif':
            return _hash_gif_no_comments(path), 'GIF without comments'
        if ft == 'rtf':
            return _hash_rtf_content(path), 'RTF without \\info'
        if ft in ('doc', 'xls', 'ppt') and _HAS_PYWIN32:
            h_core = _hash_ole_core_streams(path)
            if h_core:
                return h_core, 'Core OLE streams'
            # fallback to broader OLE content hash
            h = _hash_ole_content(path)
            if h:
                return h, 'OLE streams (excluding property sets)'
        if ft in ('docx', 'xlsx', 'pptx'):
            return _hash_office_content(path), f'{ft.upper()} content parts'
        if ft == 'word2003xml':
            # Best-effort: whole-file XML hash
            return _hash_file(path), 'XML document'
        if ft == 'pdf':
            if _HAS_PIKEPDF and pikepdf is not None:
                h = _hash_pdf_page_contents(path)
                if h:
                    return h, 'PDF page contents'
                return None, 'PDF content hash unavailable'
            else:
                return None, 'PDF hashing requires pikepdf'
        # fallback: whole-file as last resort (will change if cleaned)
        return _hash_file(path), 'Whole file'
    except Exception as e:
        return None, f'Hash error: {e}'


def _hash_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _hash_jpeg_scan(path: str) -> str:
    # Hash from SOS marker through EOI inclusive
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        if f.read(2) != b"\xff\xd8":
            return _hash_file(path)
        while True:
            marker = f.read(2)
            if len(marker) < 2:
                break
            if marker[0] != 0xFF:
                break
            if marker[1] == 0xDA:  # SOS
                # include marker itself and the rest of file up to EOF
                h.update(marker)
                while True:
                    data = f.read(1024 * 1024)
                    if not data:
                        break
                    h.update(data)
                break
            if marker[1] in (0xD8, 0xD9):
                continue
            lb = f.read(2)
            if len(lb) < 2:
                break
            length = struct.unpack('>H', lb)[0]
            f.seek(length - 2, os.SEEK_CUR)
    return h.hexdigest()


def _hash_png_idat(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        sig = f.read(8)
        if sig != b"\x89PNG\r\n\x1a\n":
            return _hash_file(path)
        while True:
            len_bytes = f.read(4)
            if len(len_bytes) < 4:
                break
            length = struct.unpack('>I', len_bytes)[0]
            ctype = f.read(4)
            data = f.read(length)
            crc = f.read(4)
            if len(ctype) < 4 or len(data) < length or len(crc) < 4:
                break
            if ctype == b'IDAT':
                h.update(data)
    return h.hexdigest()


def _hash_gif_no_comments(path: str) -> str:
    # Hash all bytes except comment extension blocks
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        hdr = f.read(6)
        if hdr not in (b'GIF87a', b'GIF89a'):
            return _hash_file(path)
        h.update(hdr)
        lsd = f.read(7)
        if len(lsd) < 7:
            return _hash_file(path)
        h.update(lsd)
        packed = lsd[4]
        gct_flag = (packed & 0x80) != 0
        gct_size = 3 * (2 ** ((packed & 0x07) + 1)) if gct_flag else 0
        if gct_size:
            h.update(f.read(gct_size))
        while True:
            introducer = f.read(1)
            if not introducer:
                break
            b = introducer[0]
            if b == 0x3B:  # trailer
                h.update(introducer)
                break
            if b == 0x2C:  # image descriptor
                h.update(introducer)
                h.update(f.read(9))
                packed2 = f.read(1)
                if not packed2:
                    break
                h.update(packed2)
                lct_flag = (packed2[0] & 0x80) != 0
                lct_size = 3 * (2 ** ((packed2[0] & 0x07) + 1)) if lct_flag else 0
                if lct_size:
                    h.update(f.read(lct_size))
                # LZW min code size
                mcs = f.read(1)
                if not mcs:
                    break
                h.update(mcs)
                # image data sub-blocks
                while True:
                    szb = f.read(1)
                    if not szb:
                        break
                    h.update(szb)
                    sz = szb[0]
                    if sz == 0:
                        break
                    h.update(f.read(sz))
                continue
            if b == 0x21:  # extension
                label = f.read(1)
                if not label:
                    break
                if label[0] == 0xFE:  # comment - skip
                    _skip_sub_blocks(f)
                    continue
                else:
                    h.update(introducer)
                    h.update(label)
                    # copy sub-blocks
                    while True:
                        szb = f.read(1)
                        if not szb:
                            break
                        h.update(szb)
                        sz = szb[0]
                        if sz == 0:
                            break
                        h.update(f.read(sz))
                    continue
            else:
                break
    return h.hexdigest()


def _hash_office_content(path: str) -> str:
    # Hash all parts except metadata and relationship control files
    names = []
    with ZipFile(path, 'r') as z:
        for info in z.infolist():
            n = info.filename
            if n.startswith('docProps/'):
                continue
            if n.endswith('.rels') or '/_rels/' in n or n.startswith('_rels/'):
                continue
            if n == '[Content_Types].xml':
                continue
            names.append(n)
        names.sort()
        h = hashlib.sha256()
        for n in names:
            h.update(n.encode('utf-8'))
            h.update(z.read(n))
    return h.hexdigest()


def _hash_pdf_page_contents(path: str) -> str | None:
    try:
        assert _HAS_PIKEPDF and pikepdf is not None
        h = hashlib.sha256()
        with pikepdf.open(path) as pdf:  # type: ignore
            for page in pdf.pages:
                contents = page.get('/Contents', None)
                if contents is None:
                    continue
                if isinstance(contents, pikepdf.Array):  # type: ignore
                    for obj in contents:
                        try:
                            h.update(bytes(obj.read_bytes()))
                        except Exception:
                            pass
                else:
                    try:
                        h.update(bytes(contents.read_bytes()))
                    except Exception:
                        pass
        return h.hexdigest()
    except Exception:
        return None


def _hash_ole_content(path: str) -> str | None:
    """Order-independent hash of all OLE streams except property sets.

    We gather (path, sha256(stream)) pairs, sort by path, then hash the
    concatenation to avoid enumeration-order differences.
    """
    try:
        stg = pythoncom.StgOpenStorage(path, None,
                                       storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE,
                                       None, 0)  # type: ignore
        entries: list[tuple[str, str]] = []

        def walk(storage, prefix: str = ""):
            try:
                enum = storage.EnumElements(0, None, 0)
            except Exception:
                return
            while True:
                try:
                    res = enum.Next(1)
                except Exception:
                    break
                if not res:
                    break
                stat = res[0]
                try:
                    name = stat[0]
                    obj_type = stat[1]
                except Exception:
                    break
                name = name or ""
                full = prefix + name
                if obj_type == storagecon.STGTY_STORAGE:  # type: ignore
                    try:
                        sub = storage.OpenStorage(name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE, None, 0)  # type: ignore
                        walk(sub, full + "/")
                    except Exception:
                        continue
                elif obj_type == storagecon.STGTY_STREAM:  # type: ignore
                    if name in ("\x05SummaryInformation", "\x05DocumentSummaryInformation"):
                        continue
                    # Hash the stream content
                    try:
                        sh = hashlib.sha256()
                        stream = storage.OpenStream(name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE, 0)  # type: ignore
                        while True:
                            chunk = stream.Read(8192)
                            if not chunk:
                                break
                            sh.update(chunk)
                        entries.append((full, sh.hexdigest()))
                    except Exception:
                        continue
                else:
                    continue

        walk(stg, "")
        entries.sort(key=lambda x: x[0])
        h = hashlib.sha256()
        for path_name, digest in entries:
            h.update(path_name.encode('utf-8', errors='ignore'))
            h.update(digest.encode('ascii'))
        return h.hexdigest()
    except Exception:
        return None


def _hash_ole_core_streams(path: str) -> str | None:
    """Hash only the main content streams for legacy Office files.

    - Word: 'WordDocument', '0Table', '1Table'
    - Excel: 'Workbook' or 'Book'
    - PowerPoint: 'PowerPoint Document'
    Ignores metadata, previews, Current User, etc. Order‑independent.
    """
    try:
        stg = pythoncom.StgOpenStorage(path, None,
                                       storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE,
                                       None, 0)  # type: ignore
        targets = {
            'worddocument', '0table', '1table',
            'workbook', 'book',
            'powerpoint document'
        }
        entries: list[tuple[str, str]] = []

        def walk(storage, prefix: str = ""):
            try:
                enum = storage.EnumElements(0, None, 0)
            except Exception:
                return
            while True:
                try:
                    res = enum.Next(1)
                except Exception:
                    break
                if not res:
                    break
                stat = res[0]
                try:
                    name = stat[0]
                    obj_type = stat[1]
                except Exception:
                    break
                name = name or ""
                full = prefix + name
                if obj_type == storagecon.STGTY_STORAGE:  # type: ignore
                    try:
                        sub = storage.OpenStorage(name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE, None, 0)  # type: ignore
                        walk(sub, full + "/")
                    except Exception:
                        continue
                elif obj_type == storagecon.STGTY_STREAM:  # type: ignore
                    base = name.lower()
                    if base in targets:
                        try:
                            sh = hashlib.sha256()
                            stream = storage.OpenStream(name, None, storagecon.STGM_READ | storagecon.STGM_SHARE_DENY_NONE, 0)  # type: ignore
                            while True:
                                chunk = stream.Read(16384)
                                if not chunk:
                                    break
                                sh.update(chunk)
                            entries.append((full, sh.hexdigest()))
                        except Exception:
                            continue
                else:
                    continue

        walk(stg, "")
        if not entries:
            return None
        entries.sort(key=lambda x: x[0])
        h = hashlib.sha256()
        for path_name, digest in entries:
            h.update(path_name.encode('utf-8', errors='ignore'))
            h.update(digest.encode('ascii'))
        return h.hexdigest()
    except Exception:
        return None
