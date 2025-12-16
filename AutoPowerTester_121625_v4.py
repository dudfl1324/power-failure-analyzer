"""AutoPowerTester_121625_v4

Refactor of AutoPowerTester_121625_v3.py with reduced repetition.
UI layout/fonts/text/messages are intentionally kept identical to v3.

Maintenance notes:
- Keep all imports (even if seemingly unused) to preserve runtime behavior.
- Avoid changing widget text, geometry, fonts, and messagebox strings.
"""

# --- Imports retained exactly as in v3 (do not remove) ---
import os
import sys
import time
import threading
import subprocess
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog


# --- Small, maintenance-focused helpers ---

def _grid(widget, **kw):
    """Thin wrapper to reduce repetitive grid() calls."""
    widget.grid(**kw)
    return widget


def _set_state(widgets, state):
    """Set state for a list/tuple of widgets."""
    for w in widgets:
        try:
            w.configure(state=state)
        except Exception:
            pass


def _safe_str(v):
    return "" if v is None else str(v)


class AutoPowerTesterApp(tk.Tk):
    """Primary UI application class (UI identical to v3)."""

    def __init__(self):
        super().__init__()

        # ---- Window (do not change title/size) ----
        self.title("Auto Power Tester")
        self.geometry("900x650")
        self.resizable(False, False)

        # ---- Runtime fields ----
        self._stop_event = threading.Event()
        self._worker = None
        self._start_time = None

        # ---- Tk variables ----
        self.var_port = tk.StringVar(value="")
        self.var_baud = tk.StringVar(value="115200")
        self.var_iterations = tk.StringVar(value="1")
        self.var_delay = tk.StringVar(value="1")
        self.var_log_dir = tk.StringVar(value="")
        self.var_script = tk.StringVar(value="")

        # ---- Fonts (keep consistent) ----
        self.font_header = ("Segoe UI", 14, "bold")
        self.font_label = ("Segoe UI", 10)
        self.font_button = ("Segoe UI", 10, "bold")
        self.font_mono = ("Consolas", 10)

        # ---- Layout ----
        self._build_ui()

        # ---- Close handler ----
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------------- UI construction ----------------
    def _build_ui(self):
        # Header
        _grid(tk.Label(self, text="Auto Power Tester", font=self.font_header), row=0, column=0, columnspan=6, pady=(12, 8))

        # Main frame
        main = _grid(tk.Frame(self), row=1, column=0, columnspan=6, sticky="nsew", padx=14)

        # Left: settings
        settings = _grid(tk.LabelFrame(main, text="Settings", font=self.font_label), row=0, column=0, sticky="nw", padx=(0, 10), pady=(0, 10))

        # Right: status/log
        status = _grid(tk.LabelFrame(main, text="Status", font=self.font_label), row=0, column=1, sticky="ne", pady=(0, 10))

        # --- Settings widgets (strings/layout must match v3) ---
        row = 0
        def add_labeled_entry(parent, label, var, width=28):
            nonlocal row
            _grid(tk.Label(parent, text=label, font=self.font_label, anchor="w"), row=row, column=0, sticky="w", padx=10, pady=6)
            ent = _grid(tk.Entry(parent, textvariable=var, width=width, font=self.font_label), row=row, column=1, sticky="w", padx=10, pady=6)
            row += 1
            return ent

        self.ent_port = add_labeled_entry(settings, "Serial Port:", self.var_port)
        self.ent_baud = add_labeled_entry(settings, "Baud Rate:", self.var_baud)
        self.ent_iterations = add_labeled_entry(settings, "Iterations:", self.var_iterations)
        self.ent_delay = add_labeled_entry(settings, "Delay (sec):", self.var_delay)

        # Log directory selector
        _grid(tk.Label(settings, text="Log Directory:", font=self.font_label, anchor="w"), row=row, column=0, sticky="w", padx=10, pady=6)
        self.ent_log_dir = _grid(tk.Entry(settings, textvariable=self.var_log_dir, width=28, font=self.font_label), row=row, column=1, sticky="w", padx=10, pady=6)
        self.btn_browse_log = _grid(tk.Button(settings, text="Browse", font=self.font_button, command=self._browse_log_dir, width=10), row=row, column=2, sticky="w", padx=(0, 10), pady=6)
        row += 1

        # Script selector
        _grid(tk.Label(settings, text="Test Script:", font=self.font_label, anchor="w"), row=row, column=0, sticky="w", padx=10, pady=6)
        self.ent_script = _grid(tk.Entry(settings, textvariable=self.var_script, width=28, font=self.font_label), row=row, column=1, sticky="w", padx=10, pady=6)
        self.btn_browse_script = _grid(tk.Button(settings, text="Browse", font=self.font_button, command=self._browse_script, width=10), row=row, column=2, sticky="w", padx=(0, 10), pady=6)
        row += 1

        # Action buttons
        btns = _grid(tk.Frame(settings), row=row, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 10))
        self.btn_start = _grid(tk.Button(btns, text="Start", font=self.font_button, width=12, command=self._start), row=0, column=0, padx=(0, 10))
        self.btn_stop = _grid(tk.Button(btns, text="Stop", font=self.font_button, width=12, command=self._stop, state="disabled"), row=0, column=1)

        # --- Status widgets ---
        srow = 0
        def add_status_line(label, var_name):
            nonlocal srow
            _grid(tk.Label(status, text=label, font=self.font_label, anchor="w"), row=srow, column=0, sticky="w", padx=10, pady=6)
            v = tk.StringVar(value="")
            setattr(self, var_name, v)
            _grid(tk.Label(status, textvariable=v, font=self.font_label, anchor="w"), row=srow, column=1, sticky="w", padx=10, pady=6)
            srow += 1

        add_status_line("State:", "var_state")
        add_status_line("Current Iteration:", "var_current_iter")
        add_status_line("Elapsed:", "var_elapsed")
        add_status_line("Last Result:", "var_last_result")

        # Log text
        self.txt_log = _grid(tk.Text(status, width=55, height=22, font=self.font_mono, wrap="none"), row=srow, column=0, columnspan=2, padx=10, pady=(10, 10))
        self.txt_log.configure(state="disabled")

        # Make columns align
        settings.grid_columnconfigure(1, weight=1)
        status.grid_columnconfigure(1, weight=1)

        # Initial state
        self._set_state_idle()

    # ---------------- UI state helpers ----------------
    def _set_state_idle(self):
        self.var_state.set("Idle")
        _set_state(
            [self.ent_port, self.ent_baud, self.ent_iterations, self.ent_delay, self.ent_log_dir, self.ent_script, self.btn_browse_log, self.btn_browse_script, self.btn_start],
            "normal",
        )
        _set_state([self.btn_stop], "disabled")

    def _set_state_running(self):
        self.var_state.set("Running")
        _set_state(
            [self.ent_port, self.ent_baud, self.ent_iterations, self.ent_delay, self.ent_log_dir, self.ent_script, self.btn_browse_log, self.btn_browse_script, self.btn_start],
            "disabled",
        )
        _set_state([self.btn_stop], "normal")

    # ---------------- Logging ----------------
    def _log(self, msg):
        """Append to UI log without changing existing text formatting in use."""
        ts = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", line)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ---------------- Browsers ----------------
    def _browse_log_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.var_log_dir.set(d)

    def _browse_script(self):
        p = filedialog.askopenfilename(filetypes=[("Python Files", "*.py"), ("All Files", "*")])
        if p:
            self.var_script.set(p)

    # ---------------- Validation (messages must match v3) ----------------
    def _validate_inputs(self):
        if not self.var_port.get().strip():
            messagebox.showerror("Error", "Serial Port is required.")
            return False
        if not self.var_baud.get().strip():
            messagebox.showerror("Error", "Baud Rate is required.")
            return False
        if not self.var_iterations.get().strip():
            messagebox.showerror("Error", "Iterations is required.")
            return False
        if not self.var_delay.get().strip():
            messagebox.showerror("Error", "Delay (sec) is required.")
            return False
        if not self.var_log_dir.get().strip():
            messagebox.showerror("Error", "Log Directory is required.")
            return False
        if not self.var_script.get().strip():
            messagebox.showerror("Error", "Test Script is required.")
            return False

        try:
            int(self.var_baud.get())
        except Exception:
            messagebox.showerror("Error", "Baud Rate must be an integer.")
            return False

        try:
            iters = int(self.var_iterations.get())
            if iters <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "Iterations must be a positive integer.")
            return False

        try:
            delay = float(self.var_delay.get())
            if delay < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "Delay (sec) must be a non-negative number.")
            return False

        if not os.path.isdir(self.var_log_dir.get()):
            messagebox.showerror("Error", "Log Directory does not exist.")
            return False

        if not os.path.isfile(self.var_script.get()):
            messagebox.showerror("Error", "Test Script does not exist.")
            return False

        return True

    # ---------------- Start/Stop ----------------
    def _start(self):
        if not self._validate_inputs():
            return

        self._stop_event.clear()
        self._start_time = time.time()
        self.var_current_iter.set("0")
        self.var_elapsed.set("00:00:00")
        self.var_last_result.set("")

        self._set_state_running()
        self._log("Starting test...")

        self._worker = threading.Thread(target=self._run_test, daemon=True)
        self._worker.start()
        self.after(250, self._update_elapsed)

    def _stop(self):
        self._stop_event.set()
        self._log("Stop requested.")

    def _update_elapsed(self):
        if self._start_time is None:
            return
        if self.var_state.get() != "Running":
            return

        elapsed = int(time.time() - self._start_time)
        h = elapsed // 3600
        m = (elapsed % 3600) // 60
        s = elapsed % 60
        self.var_elapsed.set(f"{h:02d}:{m:02d}:{s:02d}")
        self.after(250, self._update_elapsed)

    # ---------------- Worker logic (behavior preserved) ----------------
    def _run_test(self):
        port = self.var_port.get().strip()
        baud = self.var_baud.get().strip()
        iterations = int(self.var_iterations.get().strip())
        delay = float(self.var_delay.get().strip())
        log_dir = self.var_log_dir.get().strip()
        script_path = self.var_script.get().strip()

        for i in range(1, iterations + 1):
            if self._stop_event.is_set():
                self._ui_call(self._finish, "Stopped")
                return

            self._ui_call(self.var_current_iter.set, str(i))
            self._ui_call(self._log, f"Iteration {i} of {iterations}")

            ok, result_msg = self._execute_script(script_path, port, baud, log_dir, i)
            self._ui_call(self.var_last_result.set, result_msg)

            if not ok:
                self._ui_call(messagebox.showerror, "Error", result_msg)
                self._ui_call(self._finish, "Error")
                return

            if delay:
                t_end = time.time() + delay
                while time.time() < t_end:
                    if self._stop_event.is_set():
                        self._ui_call(self._finish, "Stopped")
                        return
                    time.sleep(0.05)

        self._ui_call(self._finish, "Done")

    def _execute_script(self, script_path, port, baud, log_dir, iteration):
        """Execute external script the same way as v3 (arguments preserved)."""
        try:
            log_name = f"iteration_{iteration}.log"
            log_path = os.path.join(log_dir, log_name)

            cmd = [sys.executable, script_path, "--port", port, "--baud", _safe_str(baud), "--log", log_path]
            self._ui_call(self._log, "Running: " + " ".join(cmd))

            proc = subprocess.run(cmd, capture_output=True, text=True)
            if proc.stdout:
                self._ui_call(self._log, proc.stdout.rstrip("\n"))
            if proc.stderr:
                self._ui_call(self._log, proc.stderr.rstrip("\n"))

            if proc.returncode != 0:
                return False, f"Test script failed (return code {proc.returncode})."

            return True, "Pass"
        except Exception as e:
            return False, f"Exception while running test script: {e}"

    # ---------------- UI threading ----------------
    def _ui_call(self, func, *args, **kwargs):
        self.after(0, lambda: func(*args, **kwargs))

    def _finish(self, final_state):
        self.var_state.set(final_state)
        self._set_state_idle()
        if final_state == "Done":
            self._log("Test finished.")
        elif final_state == "Stopped":
            self._log("Test stopped.")
        else:
            self._log("Test ended.")

    def _on_close(self):
        if self.var_state.get() == "Running":
            if not messagebox.askyesno("Confirm", "A test is running. Stop and exit?"):
                return
            self._stop_event.set()
        self.destroy()


def main():
    app = AutoPowerTesterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
