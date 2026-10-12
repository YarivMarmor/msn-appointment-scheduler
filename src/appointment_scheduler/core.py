from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Iterable

import numpy as np
from scipy.stats import norm


class Distribution(str, Enum):
    NORMAL = "Normal Distribution"
    UNIFORM = "Uniform Distribution"
    EMPIRICAL = "Empirical Distribution"


@dataclass(slots=True)
class SchedulerConfig:
    appointments_per_server: int = 5
    servers: int = 3
    start_hour: float = 7.0
    qos: float = 0.8
    allowed_wait_min: float = 0.0
    no_show_probability: float = 0.1
    distribution: Distribution = Distribution.NORMAL
    normal_mean: float = 30.0
    normal_sd: float = 6.0
    uniform_low: float = 10.0
    uniform_high: float = 20.0
    instances: int = 2_000
    seed: int | None = None
    empirical_values: np.ndarray | None = None

    @property
    def total_appointments(self) -> int:
        return self.appointments_per_server * self.servers

    def validate(self) -> None:
        if self.appointments_per_server < 1:
            raise ValueError("appointments_per_server must be >= 1")
        if self.servers < 1:
            raise ValueError("servers must be >= 1")
        if not 0 <= self.start_hour <= 24:
            raise ValueError("start_hour must be between 0 and 24")
        if not 0 <= self.qos < 1:
            raise ValueError("qos must be in [0, 1)")
        if self.allowed_wait_min < 0:
            raise ValueError("allowed_wait_min must be >= 0")
        if not 0 <= self.no_show_probability < 1:
            raise ValueError("no_show_probability must be in [0, 1)")
        if self.instances < 1:
            raise ValueError("instances must be >= 1")
        if self.distribution == Distribution.NORMAL:
            if self.normal_mean < 0 or self.normal_sd <= 0:
                raise ValueError("normal mean must be >= 0 and sd > 0")
        elif self.distribution == Distribution.UNIFORM:
            if self.uniform_low < 0 or self.uniform_high <= self.uniform_low:
                raise ValueError("uniform bounds must satisfy 0 <= low < high")
        elif self.distribution == Distribution.EMPIRICAL:
            if self.empirical_values is None or len(self.empirical_values) < 2:
                raise ValueError("empirical distribution requires at least 2 observations")
            if np.any(np.asarray(self.empirical_values, dtype=float) < 0):
                raise ValueError("empirical service durations must be non-negative")


@dataclass(slots=True)
class ScheduleResult:
    appointment: np.ndarray
    appointment_time_min: np.ndarray
    slot_min: np.ndarray
    expected_delay_min: np.ndarray
    service_level: np.ndarray
    avg_duration_min: np.ndarray
    distribution: str

    def as_rows(self) -> list[dict[str, float | int | str | None]]:
        rows = []
        n = len(self.appointment)
        for i in range(n):
            rows.append(
                {
                    "Appointment": int(self.appointment[i]),
                    "Time": float(self.appointment_time_min[i]),
                    "E[Slot]": None if np.isnan(self.slot_min[i]) else float(self.slot_min[i]),
                    "E[Delay]": float(self.expected_delay_min[i]),
                    "E[Service Level]": float(self.service_level[i]),
                    "avg Duration": float(self.avg_duration_min[i]),
                    "Distribution": self.distribution,
                }
            )
        return rows


def load_empirical_file(path: str | Path) -> np.ndarray:
    values: list[float] = []
    with open(path, "r", encoding="utf-8-sig") as f:
        for line_number, line in enumerate(f, start=1):
            text = line.strip()
            if not text:
                continue
            try:
                values.append(float(text))
            except ValueError as exc:
                raise ValueError(f"Invalid number on line {line_number}: {text!r}") from exc
    if len(values) < 2:
        raise ValueError("Empirical file must contain at least two numeric observations")
    return np.sort(np.asarray(values, dtype=float))


def _sample_durations(cfg: SchedulerConfig, rng: np.random.Generator) -> np.ndarray:
    n = cfg.total_appointments
    m = cfg.instances
    no_show = rng.random((n, m)) <= cfg.no_show_probability

    if cfg.distribution == Distribution.NORMAL:
        # Faithful to the VBA: negative normal draws are rejected and re-drawn.
        u = rng.random((n, m))
        x = norm.ppf(u, loc=cfg.normal_mean, scale=cfg.normal_sd)
        mask = (x < 0) & (~no_show)
        while np.any(mask):
            u2 = rng.random(np.count_nonzero(mask))
            x[mask] = norm.ppf(u2, loc=cfg.normal_mean, scale=cfg.normal_sd)
            mask = (x < 0) & (~no_show)
        x[no_show] = 0.0
        return x

    if cfg.distribution == Distribution.UNIFORM:
        u = rng.random((n, m))
        x = cfg.uniform_low + u * (cfg.uniform_high - cfg.uniform_low)
        x[no_show] = 0.0
        return x

    data = np.sort(np.asarray(cfg.empirical_values, dtype=float))
    u = rng.random((n, m))

    # Compatibility mode: mirrors the original VBA interpolation exactly.
    # The VBA multiplies U by (UBound(dataArray) - 1), which leaves the final
    # empirical observation unused for continuous U in [0,1). This is kept
    # intentionally here until the legacy behavior is validated and/or fixed.
    scale = len(data) - 2
    a = u * scale
    idx = np.floor(a).astype(int)
    idx = np.clip(idx, 0, len(data) - 2)
    frac = a - idx
    x = data[idx] + frac * (data[idx + 1] - data[idx])
    x[no_show] = 0.0
    return x


def calculate_schedule(cfg: SchedulerConfig) -> ScheduleResult:
    """Port of the original VBA appointment scheduling simulation.

    Times in the simulation are stored in minutes from midnight to make the
    algorithm independent of Excel serial-date values.
    """
    cfg.validate()
    rng = np.random.default_rng(cfg.seed)

    n = cfg.total_appointments
    k_servers = cfg.servers
    m = cfg.instances
    start_min = cfg.start_hour * 60.0

    tij = _sample_durations(cfg, rng)
    avg_duration = tij.mean(axis=1)

    ai = np.zeros(n + 1, dtype=float)
    ai[:k_servers] = start_min
    delay = np.zeros(n, dtype=float)
    service_fail = np.zeros(n, dtype=float)

    # VBA sorts the first K durations independently for each simulation instance.
    tij[:k_servers, :] = np.sort(tij[:k_servers, :], axis=0)
    first_finish = ai[0] + tij[0, :]
    fis = first_finish - cfg.allowed_wait_min
    tij[:k_servers, :] = ai[:k_servers, None] + tij[:k_servers, :]

    next_idx = k_servers
    if next_idx < n:
        ai[next_idx] = np.quantile(fis, cfg.qos, method="linear")
        actual_ready = fis + cfg.allowed_wait_min
        late = actual_ready > ai[next_idx]
        delay[next_idx] = np.where(late, actual_ready - ai[next_idx], 0.0).mean()
        service_fail[next_idx] = late.mean()

    # Mirrors VBA loop: scan = number_of_servers + 1 To number_of_appointments.
    # Python index scan corresponds to VBA scan-1.
    for scan in range(k_servers, n):
        prev_same_server = scan - k_servers
        tij[scan, :] = tij[scan, :] + np.maximum(ai[scan], tij[prev_same_server, :])

        active_start = scan - (k_servers - 1)
        active = tij[active_start : scan + 1, :]
        active.sort(axis=0)
        tij[active_start : scan + 1, :] = active

        fis = tij[active_start, :] - cfg.allowed_wait_min
        if scan + 1 < len(ai):
            ai[scan + 1] = np.quantile(fis, cfg.qos, method="linear")

        target = scan + 1
        if target < n:
            actual_ready = fis + cfg.allowed_wait_min
            late = actual_ready > ai[target]
            delay[target] = np.where(late, actual_ready - ai[target], 0.0).mean()
            service_fail[target] = late.mean()

    slot = np.full(n, np.nan, dtype=float)
    if n > 1:
        slot[:-1] = np.diff(ai[:n])

    return ScheduleResult(
        appointment=np.arange(1, n + 1),
        appointment_time_min=ai[:n],
        slot_min=slot,
        expected_delay_min=delay,
        service_level=1.0 - service_fail,
        avg_duration_min=avg_duration,
        distribution=cfg.distribution.value,
    )


def minutes_to_clock(minutes: float) -> str:
    minute = int(round(minutes)) % (24 * 60)
    return f"{minute // 60:02d}:{minute % 60:02d}"
