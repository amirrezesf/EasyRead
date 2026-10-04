"""EasyRead GUI for extracting and transcribing mixed study sources."""

import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import functools
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable, Optional

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except ImportError:  # pragma: no cover
    TKDND_AVAILABLE = False

try:
    import ttkbootstrap as tb
    from ttkbootstrap.constants import *
    TTKBOOTSTRAP_AVAILABLE = True
except ImportError:  # pragma: no cover
    TTKBOOTSTRAP_AVAILABLE = False

from source_processing import process_source
from prompt_generation import PURPOSES, write_prompt


class Tooltip:
    """Simple tooltip for a widget."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#ffffe0",
            relief=tk.SOLID,
            borderwidth=1,
            font=("tahoma", "8", "normal"),
        )
        label.pack(ipadx=1)

    def hide(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
        self.tipwindow = None


# Determine base class for TkinterDnD support
if TKDND_AVAILABLE:
    BaseClass = TkinterDnD.Tk
else:
    BaseClass = tk.Tk


class EasyReadApp(BaseClass):
    def __init__(self):
        super().__init__()
        self.title("EasyRead")
        self.minsize(640, 480)
        self.geometry("760x560")

        # Apply ttkbootstrap theme if available
        if TTKBOOTSTRAP_AVAILABLE:
            self.style = tb.Style()
            self.style.theme_use('flatly')  # you can change the theme as needed

        self.sources: list[str] = []
        self.output_dir_var = tk.StringVar()
        self.model_var = tk.StringVar(value="large-v3-turbo")
        self.language_var = tk.StringVar(value="Persian")
        self.prompt_purpose_var = tk.StringVar(value="Night-before exam handout")
        self.transcribe_var = tk.BooleanVar(value=True)
        self._busy: bool = False
        self._cancel_event = threading.Event()
        self._log_queue: queue.Queue[str] = queue.Queue()
        self._progress_var = tk.DoubleVar(value=0)

        # Session file path (in user's app data directory - XDG compliant)
        if sys.platform == "win32":
            base_dir = Path(os.getenv("APPDATA", Path.home()))
        else:
            base_dir = Path(os.getenv("XDG_DATA_HOME", Path.home() / ".local" / "share"))
        self.session_file = base_dir / "EasyRead" / "session.json"
        self.session_file.parent.mkdir(parents=True, exist_ok=True)
        self.session_artifacts_dir = self.session_file.parent / "extracted"

        self._build_ui()
        self.after(100, self._drain_log_queue)
        self._load_session()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.columnconfigure(1, weight=1)

        # Sources
        ttk.Label(frm, text="Sources").grid(row=0, column=0, sticky="nw", **pad)
        source_frame = ttk.Frame(frm)
        source_frame.grid(row=0, column=1, columnspan=2, sticky="nsew", **pad)
        source_frame.columnconfigure(0, weight=1)
        source_frame.rowconfigure(0, weight=1)

        self.source_list = tk.Listbox(source_frame, height=6, selectmode=tk.EXTENDED)
        self.source_list.grid(row=0, column=0, rowspan=2, sticky="nsew")
        source_scroll = ttk.Scrollbar(source_frame, command=self.source_list.yview)
        source_scroll.grid(row=0, column=1, rowspan=2, sticky="ns")
        self.source_list.configure(yscrollcommand=source_scroll.set)

        btn_frame = ttk.Frame(source_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, sticky="ns", padx=(8, 0))
        ttk.Button(btn_frame, text="Add files", command=self._browse_sources).grid(
            row=0, column=0, sticky="ew", pady=(0, 4)
        )
        ttk.Button(btn_frame, text="Remove selected", command=self._remove_sources).grid(
            row=1, column=0, sticky="ew"
        )
        if TTKBOOTSTRAP_AVAILABLE:
            ttk.Button(
                btn_frame, text="Clear All", command=self._clear_sources
            ).grid(row=2, column=0, sticky="ew", pady=(4, 0))

        # Enable drag-and-drop if available
        if TKDND_AVAILABLE:
            self.source_list.drop_target_register(DND_FILES)
            self.source_list.dnd_bind("<<Drop>>", self._on_drop)

        # Output folder
        ttk.Label(frm, text="Output folder").grid(row=1, column=0, sticky="w", **pad)
        out_frame = ttk.Frame(frm)
        out_frame.grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        out_frame.columnconfigure(0, weight=1)
        ttk.Entry(out_frame, textvariable=self.output_dir_var).grid(
            row=0, column=0, sticky="ew"
        )
        ttk.Button(out_frame, text="Browse…", command=self._browse_output).grid(
            row=0, column=1, sticky="e", padx=(4, 0)
        )
        ttk.Button(out_frame, text="Open", command=self._open_output_folder).grid(
            row=0, column=2, sticky="e", padx=(4, 0)
        )

        # Session buttons
        sess_frame = ttk.Frame(frm)
        sess_frame.grid(row=2, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Button(sess_frame, text="Save Session", command=self._save_session).grid(
            row=0, column=0, padx=(0, 4)
        )
        ttk.Button(sess_frame, text="Load Session", command=self._load_session).grid(
            row=0, column=1, padx=4
        )

        # Prompt generation
        prompt_frame = ttk.LabelFrame(frm, text="Prompt generation", padding=10)
        prompt_frame.grid(row=3, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(prompt_frame, text="Purpose").grid(row=0, column=0, sticky="w")
        purpose_combo = ttk.Combobox(
            prompt_frame,
            textvariable=self.prompt_purpose_var,
            state="readonly",
            width=32,
            values=list(PURPOSES),
        )
        purpose_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.prompt_btn = ttk.Button(
            prompt_frame, text="Create prompt from extracted sources", command=self._create_prompt
        )
        self.prompt_btn.grid(row=0, column=2, sticky="e", padx=(12, 0))

        # Tasks
        opts = ttk.LabelFrame(frm, text="Tasks", padding=10)
        opts.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Checkbutton(opts, text="Transcribe audio with Whisper", variable=self.transcribe_var).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(opts, text="Model").grid(row=1, column=0, sticky="w", pady=(8, 0))
        model_combo = ttk.Combobox(
            opts,
            textvariable=self.model_var,
            width=30,
            state="readonly",
            values=[
                "tiny",
                "tiny.en",
                "base",
                "base.en",
                "small",
                "small.en",
                "medium",
                "medium.en",
                "large",
                "large-v3-turbo",
            ],
        )
        model_combo.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        Tooltip(model_combo, "Whisper model size for transcription")

        ttk.Label(opts, text="Language").grid(row=2, column=0, sticky="w", pady=(8, 0))
        language_combo = ttk.Combobox(
            opts,
            textvariable=self.language_var,
            width=30,
            state="readonly",
            values=["Automatic", "Persian", "English"],
        )
        language_combo.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        Tooltip(language_combo, "Whisper transcription language; choose Automatic for language detection")

        # Progress bar
        self.progress = ttk.Progressbar(
            frm, variable=self._progress_var, maximum=100, mode="determinate"
        )
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        self.progress.grid_remove()  # hide initially

        # Run button
        self.run_btn = ttk.Button(frm, text="Run", command=self._on_run)
        self.run_btn.grid(row=6, column=0, columnspan=2, pady=(8, 4))
        
        # Cancel button (initially hidden)
        self.cancel_btn = ttk.Button(frm, text="Cancel", command=self._on_cancel, state=tk.DISABLED)
        self.cancel_btn.grid(row=6, column=2, pady=(8, 4))
        self.cancel_btn.grid_remove()

        # Status
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(frm, textvariable=self.status_var).grid(
            row=7, column=0, columnspan=3, sticky="w", **pad
        )

        # Log
        log_frame = ttk.LabelFrame(frm, text="Log", padding=6)
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", **pad)
        frm.rowconfigure(8, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=16, wrap=tk.WORD, state=tk.DISABLED)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    def _browse_sources(self):
        paths = filedialog.askopenfilenames(
            title="Select sources",
            filetypes=[
                ("Supported sources", "*.pptx *.mp3 *.wav *.m4a *.aac *.wma *.ogg *.opus *.flac *.mp4 *.pdf *.docx *.txt *.md *.html"),
                ("All files", "*.*"),
            ],
        )
        for path in paths:
            if path not in self.sources:
                self.sources.append(path)
                self.source_list.insert(tk.END, path)
        if paths and not self.output_dir_var.get():
            self.output_dir_var.set(str(Path(paths[0]).parent))

    def _remove_sources(self):
        selected = list(self.source_list.curselection())
        for index in reversed(selected):
            self.source_list.delete(index)
            del self.sources[index]

    def _clear_sources(self):
        if messagebox.askyesno("Clear All", "Remove all sources from the list?"):
            self.source_list.delete(0, tk.END)
            self.sources.clear()

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir_var.set(path)

    def _open_output_folder(self):
        out_dir = self.output_dir_var.get().strip()
        if not out_dir:
            messagebox.showinfo("Output Folder", "No output folder selected.")
            return
        out_path = Path(out_dir)
        if not out_path.exists():
            messagebox.showwarning("Output Folder", f"Folder does not exist:\n{out_dir}")
            return
        # Open folder in file explorer
        if os.name == "nt":  # Windows
            os.startfile(out_path)
        elif os.name == "posix":  # macOS/Linux
            subprocess.call(["open" if sys.platform == "darwin" else "xdg-open", out_path])

    def _save_session(self):
        data = {
            "sources": self.sources,
            "output_dir": self.output_dir_var.get(),
            "model": self.model_var.get(),
            "transcribe": self.transcribe_var.get(),
            "prompt_purpose": self.prompt_purpose_var.get(),
        }
        try:
            with open(self.session_file, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            self.log(f"Session saved to {self.session_file}")
        except Exception as e:
            messagebox.showerror("Save Session", f"Failed to save session:\n{e}")

    def _load_session(self):
        if not self.session_file.exists():
            return
        try:
            with open(self.session_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.sources = data.get("sources", [])
            self.output_dir_var.set(data.get("output_dir", ""))
            self.model_var.set(data.get("model", "large-v3-turbo"))
            self.transcribe_var.set(data.get("transcribe", True))
            self.prompt_purpose_var.set(data.get("prompt_purpose", "Night-before exam handout"))
            # Update listbox
            self.source_list.delete(0, tk.END)
            for src in self.sources:
                self.source_list.insert(tk.END, src)
            self.log(f"Session loaded from {self.session_file}")
        except Exception as e:
            self.log(f"Failed to load session: {e}")

    def _on_drop(self, event):
        """Handle drag-and-drop files onto the source list."""
        self.tk.call("tk", "scaling", 1.0)  # reset scaling
        paths = self.tk.splitlist(event.data)
        for path in paths:
            if path not in self.sources:
                self.sources.append(path)
                self.source_list.insert(tk.END, path)
        if paths and not self.output_dir_var.get():
            self.output_dir_var.set(str(Path(paths[0]).parent))

    def _create_prompt(self):
        if self._busy:
            return
        sources = [Path(s) for s in self.sources]
        out_dir = self.output_dir_var.get().strip()
        if not sources:
            messagebox.showwarning("Missing sources", "Please add and process source files first.")
            return
        if not out_dir:
            messagebox.showwarning("Missing output folder", "Choose the output folder used for source processing.")
            return
        try:
            prompt_path = write_prompt(sources, out_dir, self.prompt_purpose_var.get())
        except Exception as exc:
            messagebox.showerror("Prompt generation failed", str(exc))
            self.log(f"Prompt generation failed: {exc}")
            return
        self.log(f"Prompt created for {self.prompt_purpose_var.get()}: {prompt_path}")
        messagebox.showinfo("Prompt created", f"Prompt saved to:\n{prompt_path}")

    def _save_artifacts_to_session(self, source, output_dir, artifacts):
        """Copy generated artifacts into the persistent session workspace."""
        output_dir = Path(output_dir)
        session_source_dir = self.session_artifacts_dir
        saved = []
        for artifact in artifacts:
            artifact_path = Path(artifact)
            try:
                relative_path = artifact_path.relative_to(output_dir)
            except ValueError:
                # Audio/video inputs are returned as artifacts for transcription,
                # but the original source should not be duplicated in the session.
                continue
            destination = session_source_dir / relative_path
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(artifact_path, destination)
            saved.append(destination)
        if saved:
            self.log(f"Saved {len(saved)} extracted artifact(s) to session: {session_source_dir}")

    def log(self, message):
        self._log_queue.put(str(message))

    def _drain_log_queue(self):
        try:
            while True:
                message = self._log_queue.get_nowait()
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.after(100, self._drain_log_queue)

    def _on_run(self):
            if self._busy:
                return

            # Reset cancellation event
            self._cancel_event.clear()

            sources = [Path(s) for s in self.sources]
            out_dir = self.output_dir_var.get().strip()
            transcribe = self.transcribe_var.get()
            model_name = self.model_var.get().strip() or "large-v3-turbo"
            language = self.language_var.get().strip() or "Persian"

            if not sources:
                messagebox.showwarning("Missing sources", "Please add at least one source file.")
                return
            missing = [path for path in sources if not path.is_file()]
            if missing:
                messagebox.showerror("File not found", "Source file does not exist:\n" + "\n".join(str(m) for m in missing))
                return
            if not out_dir:
                out_dir = str(sources[0].parent)
                self.output_dir_var.set(out_dir)

            out_dir_path = Path(out_dir)
            out_dir_path.mkdir(parents=True, exist_ok=True)

            self._busy = True
            self.run_btn.configure(state=tk.DISABLED)
            self.cancel_btn.configure(state=tk.NORMAL)
            self.cancel_btn.grid()
            self.status_var.set("Working…")
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete("1.0", tk.END)
            self.log_text.configure(state=tk.DISABLED)
            self.log(f"Sources: {len(sources)}")
            self.log(f"Output folder: {out_dir}")

            # Show progress bar
            self.progress.grid()
            self._progress_var.set(0)

            def worker():
                errors = []
                try:
                    total = len(sources)
                    for i, source in enumerate(sources, start=1):
                        # Check for cancellation
                        if self._cancel_event.is_set():
                            self.log("\n--- Cancelled by user ---")
                            break

                        self.log(f"\n=== Processing {source.name} ===")
                        try:
                            artifacts = process_source(
                                source,
                                out_dir_path,
                                transcribe,
                                model_name,
                                language=language,
                                log=self.log,
                                cancel_event=self._cancel_event,
                            )
                            for artifact in artifacts:
                                self.log(f"Created: {artifact}")
                            self._save_artifacts_to_session(source, out_dir, artifacts)
                        except Exception as exc:
                            errors.append(f"{source.name}: {exc}")
                            self.log(errors[-1])
                        # Update progress
                        progress_pct = int((i / total) * 100)
                        self.after(0, functools.partial(self._progress_var.set, progress_pct))
                finally:
                    self.after(0, lambda: self._finish(errors))

            threading.Thread(target=worker, daemon=True).start()

    def _on_cancel(self):
        """Signal the worker thread to stop."""
        if self._busy:
            self._cancel_event.set()
            self.log("Cancellation requested...")
            self.cancel_btn.configure(state=tk.DISABLED)
            self.status_var.set("Cancelling...")

    def _finish(self, errors):
        self._busy = False
        self.run_btn.configure(state=tk.NORMAL)
        self.cancel_btn.configure(state=tk.DISABLED)
        self.cancel_btn.grid_remove()
        self.progress.grid_remove()  # hide progress bar
        if errors:
            self.status_var.set("Finished with errors")
            messagebox.showerror("Finished with errors", "\n\n".join(errors))
        else:
            self.status_var.set("Done")
            self.log("\nAll selected tasks finished.")
            messagebox.showinfo("Done", "All selected tasks finished.")


def main():
    app = EasyReadApp()
    app.mainloop()


if __name__ == "__main__":
    main()