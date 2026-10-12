import numpy as np

from appointment_scheduler.core import Distribution, SchedulerConfig, calculate_schedule


def test_normal_schedule_shape_and_basic_invariants():
    cfg = SchedulerConfig(instances=500, seed=42)
    result = calculate_schedule(cfg)
    assert len(result.appointment) == cfg.appointments_per_server * cfg.servers
    assert np.all(np.diff(result.appointment_time_min) >= -1e-9)
    assert np.all((result.service_level >= 0) & (result.service_level <= 1))
    assert np.all(result.expected_delay_min >= 0)


def test_uniform_distribution_runs():
    cfg = SchedulerConfig(distribution=Distribution.UNIFORM, uniform_low=10, uniform_high=20, instances=250, seed=1)
    result = calculate_schedule(cfg)
    assert np.all(result.avg_duration_min >= 0)
    assert np.nanmin(result.slot_min[:-1]) >= -1e-9


def test_empirical_distribution_runs():
    cfg = SchedulerConfig(
        distribution=Distribution.EMPIRICAL,
        empirical_values=np.array([10.0, 12.0, 15.0, 20.0, 25.0]),
        instances=250,
        seed=1,
    )
    result = calculate_schedule(cfg)
    assert len(result.appointment) == 15
