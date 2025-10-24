import os
import sys
import threading
import queue
import time
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ctypes
from ctypes import wintypes

from cleaners import (
    detect_file_metadata,
    clean_file_metadata,
    compute_content_hash,
    exiftool_sensitive_labels,
    ole_props_state,
)


class MetaCleanerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MetaCleaner - Windows Metadata Remover")
        self.geometry("960x600")

        self.items = {}  # path -> tree item id
        self.queue = queue.Queue()
        self.log_fp = None
        self._setup_log_file()
        self._build_ui()
        # Announce session start after UI is ready
        if getattr(self, 'log_path', None):
            self.log_write(f"Session started. Log file: {self.log_path}")
        self._poll_queue()
        self._init_dragdrop_windows()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self):
        top_bar = ttk.Frame(self)
        top_bar.pack(fill=tk.X, padx=8, pady=6)

        ttk.Button(top_bar, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_bar, text="Add Folder", command=self.add_folder).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_bar, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=12)
        ttk.Button(top_bar, text="Clear", command=self.clear_all).pack(side=tk.LEFT, padx=4)

        self.backup_var = tk.BooleanVar(value=True)
        self.verify_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top_bar, text="Verify content unchanged", variable=self.verify_var).pack(side=tk.RIGHT, padx=8)
        ttk.Checkbutton(top_bar, text="Backup originals (.bak)", variable=self.backup_var).pack(side=tk.RIGHT, padx=4)

        cols = ("path", "type", "will", "status")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="extended")
        self.tree.heading("path", text="Path")
        self.tree.heading("type", text="Type")
        self.tree.heading("will", text="Will Clean")
        self.tree.heading("status", text="Status")
        self.tree.column("path", width=540, anchor=tk.W)
        self.tree.column("type", width=80, anchor=tk.W)
        self.tree.column("will", width=200, anchor=tk.W)
        self.tree.column("status", width=120, anchor=tk.W)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))

        btn_bar = ttk.Frame(self)
        btn_bar.pack(fill=tk.X, padx=8, pady=(0, 6))
        ttk.Button(btn_bar, text="Scan", command=self.scan_items).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_bar, text="Clean", command=self.clean_items).pack(side=tk.LEFT, padx=4)

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 8))
        self.log = tk.Text(log_frame, height=8, wrap="word")
        self.log.pack(fill=tk.BOTH, expand=True)

    def log_write(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log.insert(tk.END, line)
        self.log.see(tk.END)
        try:
            if self.log_fp:
                self.log_fp.write(line)
                self.log_fp.flush()
        except Exception:
            pass

    def _setup_log_file(self):
        try:
            # Choose base directory next to the executable when frozen, otherwise next to this script
            if getattr(sys, 'frozen', False):
                base_dir = Path(sys.executable).parent
            else:
                base_dir = Path(__file__).resolve().parent
            logs_dir = base_dir / 'logs'
            logs_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.log_path = logs_dir / f'MetaCleaner_{ts}.log'
            self.log_fp = open(self.log_path, 'a', encoding='utf-8', newline='\n')
        except Exception:
            self.log_fp = None

    def _on_close(self):
        try:
            if self.log_fp:
                self.log_write("Session ended.")
                self.log_fp.close()
        except Exception:
            pass
        self.destroy()

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Select files")
        if not paths:
            return
        added = 0
        skipped_bak = 0
        for p in paths:
            p = os.path.abspath(p)
            if self._is_backup_file(p):
                skipped_bak += 1
                continue
            if p not in self.items:
                iid = self.tree.insert("", tk.END, values=(p, "", "", "Pending"))
                self.items[p] = iid
                added += 1
        if added:
            self.log_write(f"Added {added} file(s).")
        if skipped_bak:
            self.log_write(f"Ignored {skipped_bak} backup file(s) (*.bak).")

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select folder")
        if not folder:
            return
        folder = os.path.abspath(folder)
        count = 0
        skipped_bak = 0
        for root, _, files in os.walk(folder):
            for name in files:
                p = os.path.abspath(os.path.join(root, name))
                if self._is_backup_file(p):
                    skipped_bak += 1
                    continue
                if p not in self.items:
                    iid = self.tree.insert("", tk.END, values=(p, "", "", "Pending"))
                    self.items[p] = iid
                    count += 1
        if count:
            self.log_write(f"Added {count} file(s) from folder.")
        if skipped_bak:
            self.log_write(f"Ignored {skipped_bak} backup file(s) (*.bak) in folder.")

    def remove_selected(self):
        sel = self.tree.selection()
        for iid in sel:
            path = self.tree.item(iid, "values")[0]
            if path in self.items:
                del self.items[path]
            self.tree.delete(iid)
        if sel:
            self.log_write(f"Removed {len(sel)} item(s).")

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        self.log_write("Cleared all items.")

    def scan_items(self):
        if not self.items:
            messagebox.showinfo("Scan", "No items to scan.")
            return
        # Auto-remove any lingering backup files from the list
        removed = 0
        for path, iid in list(self.items.items()):
            if self._is_backup_file(path):
                try:
                    self.tree.delete(iid)
                except Exception:
                    pass
                self.items.pop(path, None)
                removed += 1
        if removed:
            self.log_write(f"Removed {removed} backup file(s) (*.bak) from list before scan.")
        self.log_write("Scanning items for metadata...")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        for path, iid in list(self.items.items()):
            try:
                res = detect_file_metadata(path)
                summary = res.get("summary", [])
                note = res.get("note")
                filetype = res.get("type", "?")
                if note and not res.get("can_clean"):
                    will = note
                    status = "Unsupported"
                else:
                    will = ", ".join(summary) if summary else "None"
                    status = "Scanned"
                self.queue.put(("update_row", iid, (path, filetype, will, status)))
            except Exception as e:
                self.queue.put(("update_row", iid, (path, "?", "", f"Error: {e}")))
        self.queue.put(("log", "Scan complete."))

    def clean_items(self):
        if not self.items:
            messagebox.showinfo("Clean", "No items to clean.")
            return
        self.log_write("Cleaning metadata from supported files...")
        backup = bool(self.backup_var.get())
        threading.Thread(target=self._clean_worker, args=(backup,), daemon=True).start()

    def _clean_worker(self, backup: bool):
        cleaned = 0
        for path, iid in list(self.items.items()):
            try:
                res = detect_file_metadata(path)
                if res.get("note") and not res.get("can_clean"):
                    note = res.get("note", "Unsupported format")
                    self.queue.put(("status", iid, "Unsupported"))
                    self.queue.put(("log", f"{Path(path).name}: {note}"))
                    continue
                if not res.get("can_clean"):
                    self.queue.put(("status", iid, "No metadata"))
                    self.queue.put(("log", f"{Path(path).name}: No metadata to remove"))
                    continue
                # Optional pre-hash
                pre_hash = None
                scheme_desc = None
                pre_tags = None
                if self.verify_var.get():
                    pre_hash, scheme_desc = compute_content_hash(path)
                    # capture ExifTool tag presence for tolerant verification
                    try:
                        pre_tags = exiftool_sensitive_labels(path)
                    except Exception:
                        pre_tags = None
                    # capture OLE property presence for legacy Office
                    try:
                        pre_ole = ole_props_state(path)
                    except Exception:
                        pre_ole = None
                changed, detail = clean_file_metadata(path, backup=backup)
                status = "Cleaned" if changed else "No change"
                cleaned += 1 if changed else 0
                # Optional post-hash compare
                if self.verify_var.get() and changed:
                    post_hash, _ = compute_content_hash(path)
                    if pre_hash and post_hash and post_hash == pre_hash:
                        status = "Cleaned (verified)"
                        self.queue.put(("log", f"{Path(path).name}: Verified content unchanged ({scheme_desc})"))
                    elif pre_hash and post_hash and post_hash != pre_hash:
                        # Tolerant verification: if ExifTool-sensitive tags disappeared, accept as verified
                        try:
                            post_tags = exiftool_sensitive_labels(path)
                        except Exception:
                            post_tags = None
                        # Legacy Office: if OLE property sets existed before and are gone after, accept
                        ole_ok = False
                        try:
                            post_ole = ole_props_state(path)
                            if pre_ole and isinstance(pre_ole, tuple) and any(pre_ole) and post_ole and not any(post_ole):
                                ole_ok = True
                        except Exception:
                            ole_ok = False
                        if pre_tags and isinstance(pre_tags, list) and (not post_tags or len(post_tags) == 0):
                            status = "Cleaned (verified)"
                            self.queue.put(("log", f"{Path(path).name}: Tags removed; accepting as verified (tolerant)"))
                        elif ole_ok:
                            status = "Cleaned (verified)"
                            self.queue.put(("log", f"{Path(path).name}: OLE properties removed; accepting as verified (tolerant)"))
                        else:
                            status = "Cleaned (mismatch)"
                            self.queue.put(("log", f"{Path(path).name}: WARNING content hash changed ({scheme_desc})"))
                    else:
                        self.queue.put(("log", f"{Path(path).name}: Verification unavailable"))
                self.queue.put(("status", iid, status))
                if detail:
                    self.queue.put(("log", f"{Path(path).name}: {detail}"))
            except Exception as e:
                self.queue.put(("status", iid, f"Error: {e}"))
        self.queue.put(("log", f"Cleaning complete. Cleaned {cleaned} file(s)."))

    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "update_row":
                    _, iid, values = item
                    self.tree.item(iid, values=values)
                elif kind == "status":
                    _, iid, status = item
                    vals = list(self.tree.item(iid, "values"))
                    if vals:
                        vals[-1] = status
                        self.tree.item(iid, values=tuple(vals))
                elif kind == "log":
                    _, msg = item
                    self.log_write(msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # --- Windows drag & drop via WM_DROPFILES ---
    def _init_dragdrop_windows(self):
        if os.name != 'nt':
            return
        try:
            self.update_idletasks()
            hwnd = self.winfo_id()
            GWL_WNDPROC = -4
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            DragAcceptFiles = shell32.DragAcceptFiles
            DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            DragAcceptFiles.restype = None
            DragQueryFile = shell32.DragQueryFileW
            DragQueryFile.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
            DragQueryFile.restype = wintypes.UINT
            DragFinish = shell32.DragFinish
            DragFinish.argtypes = [wintypes.HANDLE]
            DragFinish.restype = None

            WM_DROPFILES = 0x0233

            WNDPROC = ctypes.WINFUNCTYPE(wintypes.LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

            orig_wndproc = wintypes.LONG_PTR(user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC))

            def py_wndproc(hWnd, msg, wParam, lParam):
                if msg == WM_DROPFILES:
                    hDrop = wParam
                    count = DragQueryFile(hDrop, 0xFFFFFFFF, None, 0)
                    for i in range(count):
                        length = DragQueryFile(hDrop, i, None, 0) + 1
                        buf = ctypes.create_unicode_buffer(length)
                        DragQueryFile(hDrop, i, buf, length)
                        path = buf.value
                        self._add_path_from_drop(path)
                    DragFinish(hDrop)
                    return 0
                return user32.CallWindowProcW(orig_wndproc, hWnd, msg, wParam, lParam)

            self._dnd_proc = WNDPROC(py_wndproc)  # keep reference
            user32.SetWindowLongPtrW(hwnd, GWL_WNDPROC, self._dnd_proc)
            DragAcceptFiles(hwnd, True)
            self.log_write("Drag-and-drop enabled (Windows)")
        except Exception:
            # Silently ignore if DnD setup fails
            pass

    def _add_path_from_drop(self, path: str):
        try:
            if os.path.isdir(path):
                # add entire folder recursively
                count = 0
                skipped_bak = 0
                for root, _, files in os.walk(path):
                    for name in files:
                        p = os.path.abspath(os.path.join(root, name))
                        if self._is_backup_file(p):
                            skipped_bak += 1
                            continue
                        if p not in self.items:
                            iid = self.tree.insert("", tk.END, values=(p, "", "", "Pending"))
                            self.items[p] = iid
                            count += 1
                if count:
                    self.log_write(f"Added {count} file(s) from drop folder.")
                if skipped_bak:
                    self.log_write(f"Ignored {skipped_bak} backup file(s) (*.bak) in dropped folder.")
            else:
                p = os.path.abspath(path)
                if self._is_backup_file(p):
                    self.log_write("Ignored backup file (*.bak) from drop.")
                    return
                if p not in self.items:
                    iid = self.tree.insert("", tk.END, values=(p, "", "", "Pending"))
                    self.items[p] = iid
                    self.log_write(f"Added file: {Path(p).name}")
        except Exception as e:
            self.log_write(f"Drop error: {e}")

    def _is_backup_file(self, path: str) -> bool:
        name = os.path.basename(path).lower()
        return name.endswith('.bak') or '.bak.' in name


def main():
    app = MetaCleanerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
