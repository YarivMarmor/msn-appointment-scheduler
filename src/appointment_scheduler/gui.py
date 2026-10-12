from __future__ import annotations

import csv
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from .core import (
    Distribution,
    SchedulerConfig,
    calculate_schedule,
    load_empirical_file,
    minutes_to_clock,
)

RUN_MODES = {
    "Fast for exploration": 2_000,
    "Medium - for building trust": 10_000,
    "Extensive - for final result": 50_000,
}


class SchedulerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MSN Appointment Scheduler")
        self.geometry("1180x760")
        self.minsize(1000, 680)
        self.empirical_path: Path | None = None
        self.result = None
        self._build_ui()
        self._distribution_changed()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        left = ttk.LabelFrame(outer, text="Simulation parameters", padding=12)
        left.pack(side="left", fill="y", padx=(0, 12))
        right = ttk.Frame(outer)
        right.pack(side="right", fill="both", expand=True)

        self.vars = {
            "distribution": tk.StringVar(value="Normal"),
            "normal_mean": tk.StringVar(value="30"),
            "normal_sd": tk.StringVar(value="6"),
            "uniform_low": tk.StringVar(value="10"),
            "uniform_high": tk.StringVar(value="20"),
            "appointments": tk.StringVar(value="5"),
            "servers": tk.StringVar(value="3"),
            "start": tk.StringVar(value="7"),
            "qos": tk.StringVar(value="0.8"),
            "wait": tk.StringVar(value="0"),
            "noshow": tk.StringVar(value="0.1"),
            "mode": tk.StringVar(value="Fast for exploration"),
        }

        row = 0
        ttk.Label(left, text="Distribution").grid(row=row, column=0, sticky="w", pady=4)
        combo = ttk.Combobox(
            left,
            textvariable=self.vars["distribution"],
            values=["Normal", "Uniform", "Empirical"],
            state="readonly",
            width=25,
        )
        combo.grid(row=row, column=1, sticky="ew", pady=4)
        combo.bind("<<ComboboxSelected>>", lambda _e: self._distribution_changed())
        row += 1

        self.param1_label = ttk.Label(left, text="Mean [min]")
        self.param1_label.grid(row=row, column=0, sticky="w", pady=4)
        self.param1 = ttk.Entry(left, textvariable=self.vars["normal_mean"])
        self.param1.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        self.param2_label = ttk.Label(left, text="SD [min]")
        self.param2_label.grid(row=row, column=0, sticky="w", pady=4)
        self.param2 = ttk.Entry(left, textvariable=self.vars["normal_sd"])
        self.param2.grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        self.file_button = ttk.Button(left, text="Load empirical TXT...", command=self._load_empirical)
        self.file_label = ttk.Label(left, text="No empirical file selected", wraplength=220)

        fields = [
            ("Appointments per server", "appointments"),
            ("Number of servers", "servers"),
            ("Start of day [hour]", "start"),
            ("QoS P(wait ≤ T)", "qos"),
            ("Allowed waiting time T [min]", "wait"),
            ("Probability of no-show", "noshow"),
        ]
        for label, key in fields:
            ttk.Label(left, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(left, textvariable=self.vars[key]).grid(row=row, column=1, sticky="ew", pady=4)
            row += 1

        ttk.Label(left, text="Running mode").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Combobox(
            left,
            textvariable=self.vars["mode"],
            values=list(RUN_MODES),
            state="readonly",
            width=25,
        ).grid(row=row, column=1, sticky="ew", pady=4)
        row += 1

        ttk.Button(left, text="Calculate", command=self._calculate).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(14, 5)
        )
        row += 1
        self.export_button = ttk.Button(left, text="Export results to CSV", command=self._export_csv, state="disabled")
        self.export_button.grid(row=row, column=0, columnspan=2, sticky="ew", pady=5)
        left.columnconfigure(1, weight=1)

        table_frame = ttk.LabelFrame(right, text="Appointment table", padding=8)
        table_frame.pack(fill="both", expand=True)
        columns = ("appointment", "time", "slot", "delay", "service", "duration")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        headings = {
            "appointment": "Appointment",
            "time": "Time",
            "slot": "E[Slot]",
            "delay": "E[Delay]",
            "service": "E[Service Level]",
            "duration": "Avg Duration",
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, anchor="center", width=115)
        self.tree.pack(fill="both", expand=True)

        chart_frame = ttk.LabelFrame(right, text="Results", padding=8)
        chart_frame.pack(fill="both", expand=True, pady=(12, 0))
        self.figure = Figure(figsize=(7, 3.1), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

    def _distribution_changed(self) -> None:
        dist = self.vars["distribution"].get()
        self.file_button.grid_forget()
        self.file_label.grid_forget()
        if dist == "Normal":
            self.param1_label.config(text="Mean μ [min]")
            self.param2_label.config(text="SD σ [min]")
            self.param1.config(textvariable=self.vars["normal_mean"])
            self.param2.config(textvariable=self.vars["normal_sd"])
            self.param1_label.grid()
            self.param2_label.grid()
            self.param1.grid()
            self.param2.grid()
        elif dist == "Uniform":
            self.param1_label.config(text="Lower bound a [min]")
            self.param2_label.config(text="Upper bound b [min]")
            self.param1.config(textvariable=self.vars["uniform_low"])
            self.param2.config(textvariable=self.vars["uniform_high"])
            self.param1_label.grid()
            self.param2_label.grid()
            self.param1.grid()
            self.param2.grid()
        else:
            self.param1_label.grid_remove()
            self.param2_label.grid_remove()
            self.param1.grid_remove()
            self.param2.grid_remove()
            info = self.param1_label.grid_info()
            base_row = int(info.get("row", 1)) if info else 1
            self.file_button.grid(row=base_row, column=0, columnspan=2, sticky="ew", pady=4)
            self.file_label.grid(row=base_row + 1, column=0, columnspan=2, sticky="w", pady=(0, 8))

    def _load_empirical(self) -> None:
        filename = filedialog.askopenfilename(title="Select empirical duration file", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if filename:
            try:
                values = load_empirical_file(filename)
            except Exception as exc:
                messagebox.showerror("Invalid file", str(exc))
                return
            self.empirical_path = Path(filename)
            self.file_label.config(text=f"{self.empirical_path.name} | n={len(values)} | mean={values.mean():.3f} | sd={values.std(ddof=1):.3f}")

    def _make_config(self) -> SchedulerConfig:
        name = self.vars["distribution"].get()
        dist = {"Normal": Distribution.NORMAL, "Uniform": Distribution.UNIFORM, "Empirical": Distribution.EMPIRICAL}[name]
        empirical = None
        if dist == Distribution.EMPIRICAL:
            if self.empirical_path is None:
                raise ValueError("Please select an empirical TXT file")
            empirical = load_empirical_file(self.empirical_path)
        return SchedulerConfig(
            appointments_per_server=int(float(self.vars["appointments"].get())),
            servers=int(float(self.vars["servers"].get())),
            start_hour=float(self.vars["start"].get()),
            qos=float(self.vars["qos"].get()),
            allowed_wait_min=float(self.vars["wait"].get()),
            no_show_probability=float(self.vars["noshow"].get()),
            distribution=dist,
            normal_mean=float(self.vars["normal_mean"].get()),
            normal_sd=float(self.vars["normal_sd"].get()),
            uniform_low=float(self.vars["uniform_low"].get()),
            uniform_high=float(self.vars["uniform_high"].get()),
            instances=RUN_MODES[self.vars["mode"].get()],
            empirical_values=empirical,
        )

    def _calculate(self) -> None:
        try:
            cfg = self._make_config()
            self.result = calculate_schedule(cfg)
        except Exception as exc:
            messagebox.showerror("Calculation error", str(exc))
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.result.as_rows():
            self.tree.insert(
                "",
                "end",
                values=(
                    row["Appointment"],
                    minutes_to_clock(row["Time"]),
                    "" if row["E[Slot]"] is None else f"{row['E[Slot]']:.2f}",
                    f"{row['E[Delay]']:.2f}",
                    f"{row['E[Service Level]']:.4f}",
                    f"{row['avg Duration']:.2f}",
                ),
            )

        self.ax.clear()
        self.ax.plot(self.result.appointment, self.result.appointment_time_min / 60.0, marker="o")
        self.ax.set_xlabel("Appointment")
        self.ax.set_ylabel("Time of day [hour]")
        self.ax.grid(True, alpha=0.25)
        self.figure.tight_layout()
        self.canvas.draw_idle()
        self.export_button.config(state="normal")

    def _export_csv(self) -> None:
        if self.result is None:
            return
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not filename:
            return
        rows = self.result.as_rows()
        with open(filename, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            writer.writeheader()
            for row in rows:
                row = dict(row)
                row["Time"] = minutes_to_clock(float(row["Time"]))
                writer.writerow(row)
        messagebox.showinfo("Export complete", f"Saved to:\n{filename}")


def main() -> None:
    SchedulerApp().mainloop()


if __name__ == "__main__":
    main()
