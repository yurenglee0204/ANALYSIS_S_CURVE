#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""分析 X 軸 Excel 時序資料，輸出含 α–β 濾波與 S 曲線的新活頁簿。

安裝：python -m pip install numpy scipy openpyxl xlsxwriter
視窗：python axis_scurve_analysis.py
命令列：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx"
自調適：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx" --auto-alpha-beta
卡曼濾波：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx" --kalman-filter
卡曼參數：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx" --kalman-filter --kalman-q 0.001 --kalman-r 0.5
自適應卡曼：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx" --kalman-filter --adaptive-kalman
NC Code：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx" --nc-code-file "C:\\量測資料\\program.nc"
貼上路徑：python axis_scurve_analysis.py --no-dialog
路徑：python axis_scurve_analysis.py "C:\\量測資料\\11_12_log_report.xlsx"
參數：python axis_scurve_analysis.py --help
需要 Python 3.10+。主程式接受 .xlsx / .xlsm，輸出 .xlsx。
"""

from __future__ import annotations

import argparse
import base64
from dataclasses import dataclass, replace
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unicodedata
from zipfile import BadZipFile

try:
    import numpy as np
    import openpyxl
    from scipy.optimize import least_squares
    import xlsxwriter
except ImportError as exc:
    raise SystemExit(
        f"缺少套件：{exc.name}\n請執行 python -m pip install numpy scipy openpyxl xlsxwriter"
    ) from None


INPUT_COLUMNS = ("index", "Xcmd", "Xfb", "Vxcmd", "Vxfb", "Following_error", "tick")
SIGNALS = ("Vxcmd", "Vxfb", "α–β 速度")
EXCEL_ROWS = 1_048_576
APP_VERSION = "v2026.11.17.1806"
DEFAULT_FONT_SIZE = 12


@dataclass(frozen=True)
class Settings:
    tick_ms: float
    alpha: float
    beta: float
    half_window: int
    velocity_unit: str
    motion_threshold: float
    min_motion: int
    merge_gap: int
    crossing_samples: int
    chart_ramps: int


@dataclass
class Log:
    sheet: str
    source_rows: np.ndarray
    data: np.ndarray
    speed_cmd: np.ndarray
    speed_fb: np.ndarray
    following: np.ndarray
    time: np.ndarray
    interval: np.ndarray
    gaps: np.ndarray
    filled: tuple[int, int, int]


@dataclass
class Estimate:
    predicted: np.ndarray
    residual: np.ndarray
    position: np.ndarray
    velocity: np.ndarray
    acceleration: np.ndarray
    command_acceleration: np.ndarray


@dataclass
class KalmanEstimate:
    position: np.ndarray
    velocity: np.ndarray
    acceleration: np.ndarray
    process_noise: float
    measurement_variance: float
    adaptive: bool = False
    process_noise_values: np.ndarray | None = None
    measurement_variance_values: np.ndarray | None = None


@dataclass
class Movement:
    number: int
    first: int
    stop: int
    direction: int
    block_start: int
    block_stop: int
    floor: int = 0
    ceiling: int = 0
    plateau: float | None = None
    flat: tuple[int, int] | None = None


@dataclass
class Crossing:
    before: int
    after: int
    speed_before: float
    speed_after: float
    time: float


@dataclass
class Fit:
    start: float
    rise: float
    hold: float
    fall: float
    rmse: float
    peak: float

    @property
    def boundaries(self):
        return (
            self.start,
            self.start + self.rise,
            self.start + self.rise + self.hold,
            self.start + self.rise + self.hold + self.fall,
        )


@dataclass
class Ramp:
    move: Movement
    accelerating: bool
    anchor: int
    fit: Fit | None
    status: str
    fractions: tuple[float, float, float]
    crossings: list[list[Crossing | None]]
    near_zero: Crossing | None = None
    peak_feedback: float | None = None

    @property
    def label(self):
        return "加速" if self.accelerating else "減速"


def normalize_header(value):
    text = unicodedata.normalize("NFKC", str(value or "")).casefold().lstrip("\ufeff")
    text = re.sub(r"\s+", "", text)
    return "index" if text == "index(ms)" else text


def find_input_sheet(book, name):
    if name and name not in book.sheetnames:
        raise ValueError(f"找不到工作表「{name}」。可用工作表：{', '.join(book.sheetnames)}")
    candidates = []
    for sheet in ([book[name]] if name else book.worksheets):
        # 有些記錄器寫入的工作表尺寸不正確。限制表頭搜尋範圍並重新掃描實際資料列。
        sheet.reset_dimensions()
        required = [normalize_header(h) for h in INPUT_COLUMNS]
        for row_id, row in enumerate(sheet.iter_rows(max_row=50, values_only=True), 1):
            names = [normalize_header(value) for value in row]
            if all(key in names for key in required):
                duplicated = [key for key in required if names.count(key) != 1]
                if duplicated:
                    raise ValueError(f"工作表「{sheet.title}」第 {row_id} 列有重複欄名：{', '.join(duplicated)}")
                candidates.append((sheet, row_id, [names.index(key) for key in required]))
                break
    if not candidates:
        raise ValueError(f"工作表前 50 列找不到必要欄位：{', '.join(INPUT_COLUMNS)}")
    if len(candidates) > 1:
        matching = ", ".join(sheet.title for sheet, _, _ in candidates)
        raise ValueError(f"多張工作表同時符合：{matching}。請加 --sheet 指定工作表。")
    return candidates[0]


def as_number(value, column, row, optional):
    if value is None or (isinstance(value, str) and value.strip() in {"", "#N/A", "#VALUE!", "#DIV/0!", "#REF!"}):
        if optional:
            return np.nan
        raise ValueError(f"Excel 第 {row} 列的 {column} 缺少有效數值。公式欄可先在 Excel 重新計算並存檔。")
    try:
        if isinstance(value, bool):
            raise ValueError
        number = float(value)
        if not math.isfinite(number):
            raise ValueError
    except (TypeError, ValueError, OverflowError):
        raise ValueError(f"Excel 第 {row} 列的 {column} 須為數字，讀到 {value!r}。") from None
    return number


def load_log(path, sheet_name, settings):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        sheet, header_row, columns = find_input_sheet(book, sheet_name)
        source_rows, records = [], []
        for row_id, row in enumerate(
            sheet.iter_rows(min_row=header_row + 1, max_col=max(columns) + 1, values_only=True),
            header_row + 1,
        ):
            values = [row[i] for i in columns]
            if all(value is None or (isinstance(value, str) and not value.strip()) for value in values):
                continue
            source_rows.append(row_id)
            records.append([as_number(value, INPUT_COLUMNS[i], row_id, i in (3, 4, 5))
                            for i, value in enumerate(values)])
        sheet_title = sheet.title
    finally:
        book.close()
    if len(records) < 2:
        raise ValueError("有效資料少於兩筆，無法計算速度。")

    data = np.asarray(records, dtype=np.float64)
    ticks = data[:, 6]
    step = np.diff(ticks)
    invalid = np.flatnonzero(step <= 0)
    if invalid.size:
        row = source_rows[invalid[0] + 1]
        raise ValueError(f"Excel 第 {row} 列的 tick 重複或倒退，請檢查時間排序或計數器回繞。")
    if not np.all(np.equal(ticks, np.rint(ticks))):
        raise ValueError("tick 須為整數計數。每單位計數的長度由 --tick-ms 設定。")

    time = (ticks - ticks[0]) * settings.tick_ms
    dt = np.r_[settings.tick_ms, step * settings.tick_ms]
    gaps = np.flatnonzero(step != 1) + 1
    result = []
    filled = []
    for raw_col, positions in ((3, data[:, 1]), (4, data[:, 2])):
        velocity = data[:, raw_col].copy()
        if settings.velocity_unit == "pulse/sample":
            velocity /= dt
        missing = np.isnan(velocity)
        # 若 log 由途中開始，缺少的第一筆速度以第一段量測位置差初始化。
        derivative = np.diff(positions) / dt[1:]
        velocity[missing] = np.r_[derivative[0], derivative][missing]
        filled.append(int(missing.sum()))
        result.append(velocity)
    following = data[:, 5].copy()
    missing = np.isnan(following)
    following[missing] = (data[:, 1] - data[:, 2])[missing]
    filled.append(int(missing.sum()))
    return Log(sheet_title, np.array(source_rows), data, *result, following,
               time, dt, gaps, tuple(filled))


def contiguous_blocks(length, gaps):
    cut = [0, *map(int, gaps), length]
    return list(zip(cut[:-1], cut[1:]))


def centered_difference(values, time, half, blocks):
    derivative = np.full(len(values), np.nan)
    for start, stop in blocks:
        if stop - start > 2 * half:
            derivative[start + half:stop - half] = (
                (values[start + 2 * half:stop] - values[start:stop - 2 * half])
                / (time[start + 2 * half:stop] - time[start:stop - 2 * half])
            )
    return derivative


def alpha_beta(log, settings):
    x = log.data[:, 2]
    predicted = np.full(len(x), np.nan)
    residual = np.full(len(x), np.nan)
    position = np.empty(len(x))
    velocity = np.empty(len(x))
    blocks = contiguous_blocks(len(x), log.gaps)
    for start, stop in blocks:
        position[start], velocity[start] = x[start], log.speed_fb[start]
        for k in range(start + 1, stop):
            dt = log.interval[k]
            predicted[k] = position[k - 1] + dt * velocity[k - 1]
            residual[k] = x[k] - predicted[k]
            position[k] = predicted[k] + settings.alpha * residual[k]
            velocity[k] = velocity[k - 1] + settings.beta * residual[k] / dt

    acceleration = centered_difference(velocity, log.time, settings.half_window, blocks)
    command_acc = centered_difference(log.speed_cmd, log.time, settings.half_window, blocks)
    return Estimate(predicted, residual, position, velocity, acceleration, command_acc)


def kalman_noise_defaults(log):
    observations = log.data[:, 2]
    blocks = contiguous_blocks(len(observations), log.gaps)
    sample_intervals = log.interval[1:]
    reference_dt = float(np.median(sample_intervals[sample_intervals > 0]))
    second_differences = [
        np.diff(observations[start:stop], n=2)
        for start, stop in blocks if stop - start > 2
    ]
    if second_differences:
        second_difference = np.concatenate(second_differences)
        center = np.median(second_difference)
        robust_sigma = 1.4826 * np.median(np.abs(second_difference - center)) / math.sqrt(6)
        measurement_variance = max(
            float(robust_sigma ** 2),
            np.finfo(float).eps * max(float(np.ptp(observations)) ** 2, 1.0),
        )
    else:
        measurement_variance = 1.0
    return measurement_variance / reference_dt ** 5, measurement_variance


def kalman_filter(log, process_noise=None, measurement_variance=None,
                  adaptive=False, motion_threshold=5.0):
    observations = log.data[:, 2]
    length = len(observations)
    position = np.empty(length)
    velocity = np.empty(length)
    acceleration = np.empty(length)
    sample_intervals = log.interval[1:]
    reference_dt = float(np.median(sample_intervals[sample_intervals > 0]))
    default_process_noise, default_measurement_variance = kalman_noise_defaults(log)
    process_noise = default_process_noise if process_noise is None else process_noise
    measurement_variance = (
        default_measurement_variance if measurement_variance is None else measurement_variance
    )
    if not math.isfinite(process_noise) or process_noise < 0:
        raise ValueError("卡曼濾波 Q 須為有限的非負數值。")
    if not math.isfinite(measurement_variance) or measurement_variance <= 0:
        raise ValueError("卡曼濾波 R 須為有限的正數。")
    if not math.isfinite(motion_threshold) or motion_threshold < 0:
        raise ValueError("自適應卡曼的運動速度門檻須為有限的非負數。")
    observation = np.array([[1.0, 0.0, 0.0]])
    identity = np.eye(3)
    process_noise_values = np.empty(length) if adaptive else None
    measurement_variance_values = np.empty(length) if adaptive else None
    lower_measurement_variance = max(
        measurement_variance * 0.1, np.finfo(float).tiny
    )
    upper_measurement_variance = max(
        lower_measurement_variance,
        min(measurement_variance * 10.0, np.finfo(float).max),
    )
    adaptation_rate = 0.05

    for start, stop in contiguous_blocks(length, log.gaps):
        state = np.array([observations[start], log.speed_fb[start], 0.0])
        current_measurement_variance = measurement_variance
        current_process_noise = process_noise
        if adaptive:
            is_moving = abs(log.speed_cmd[start]) > motion_threshold
            current_process_noise = process_noise if is_moving else process_noise * 0.1
            process_noise_values[start] = current_process_noise
            measurement_variance_values[start] = current_measurement_variance
        covariance = np.diag((
            measurement_variance,
            measurement_variance / reference_dt ** 2,
            measurement_variance / reference_dt ** 4,
        ))
        position[start], velocity[start], acceleration[start] = state

        for index in range(start + 1, stop):
            dt = float(log.interval[index])
            dt2, dt3, dt4, dt5 = dt ** 2, dt ** 3, dt ** 4, dt ** 5
            if adaptive:
                is_moving = abs(log.speed_cmd[index]) > motion_threshold
                current_process_noise = process_noise if is_moving else process_noise * 0.1
                process_noise_values[index] = current_process_noise
                measurement_variance_values[index] = current_measurement_variance
            transition = np.array([
                [1.0, dt, dt2 / 2],
                [0.0, 1.0, dt],
                [0.0, 0.0, 1.0],
            ])
            process_covariance = current_process_noise * np.array([
                [dt5 / 20, dt4 / 8, dt3 / 6],
                [dt4 / 8, dt3 / 3, dt2 / 2],
                [dt3 / 6, dt2 / 2, dt],
            ])
            predicted_state = transition @ state
            predicted_covariance = transition @ covariance @ transition.T + process_covariance
            if not np.isfinite(predicted_covariance).all():
                raise ValueError("卡曼濾波的 Q、R 或 tick 間隔造成數值溢位，請調整參數。")
            innovation = observations[index] - (observation @ predicted_state)[0]
            predicted_measurement_variance = predicted_covariance[0, 0]
            innovation_variance = (
                predicted_measurement_variance + current_measurement_variance
            )
            if not math.isfinite(innovation_variance) or innovation_variance <= 0:
                raise ValueError("卡曼濾波的創新變異數無效，請檢查 Q、R 參數。")
            gain = (predicted_covariance @ observation.T) / innovation_variance
            state = predicted_state + gain[:, 0] * innovation
            if not np.isfinite(state).all():
                raise ValueError("卡曼濾波估值溢位，請調小 Q 或檢查輸入資料。")
            correction = identity - gain @ observation
            covariance = (
                correction @ predicted_covariance @ correction.T
                + current_measurement_variance * (gain @ gain.T)
            )
            position[index], velocity[index], acceleration[index] = state
            if adaptive:
                innovation_measurement_variance = max(
                    innovation * innovation - predicted_measurement_variance,
                    lower_measurement_variance,
                )
                current_measurement_variance = min(
                    max(
                        (1 - adaptation_rate) * current_measurement_variance
                        + adaptation_rate * innovation_measurement_variance,
                        lower_measurement_variance,
                    ),
                    upper_measurement_variance,
                )

    return KalmanEstimate(
        position, velocity, acceleration, process_noise, measurement_variance,
        adaptive, process_noise_values, measurement_variance_values,
    )


def tune_alpha_beta(log, settings):
    sample_limit = 4096
    window_size = 256
    blocks = contiguous_blocks(len(log.time), log.gaps)
    if len(log.time) <= sample_limit:
        windows = blocks
    else:
        boundaries = np.r_[0, log.gaps, len(log.time)]
        targets = np.linspace(0, len(log.time) - 1, sample_limit // window_size, dtype=int)
        windows = []
        for target in targets:
            block_id = int(np.searchsorted(log.gaps, target, side="right"))
            block_start, block_stop = int(boundaries[block_id]), int(boundaries[block_id + 1])
            start = max(block_start, min(int(target) - window_size // 2, block_stop - window_size))
            stop = min(start + window_size, block_stop)
            if stop - start > 1 and (start, stop) not in windows:
                windows.append((start, stop))
        if not windows:
            windows = blocks

    positions = log.data[:, 2]
    feedback = log.speed_fb

    def scores(alpha_values, beta_fractions):
        alpha_values, beta_fractions = np.meshgrid(alpha_values, beta_fractions, indexing="ij")
        alphas = alpha_values.ravel()
        fractions = beta_fractions.ravel()
        betas = fractions * (4 - 2 * alphas)
        squared_errors = np.zeros(len(alphas))
        count = 0
        for start, stop in windows:
            estimated_position = float(positions[start])
            estimated_velocity = np.full(len(alphas), feedback[start])
            for index in range(start + 1, stop):
                dt = log.interval[index]
                predicted = estimated_position + dt * estimated_velocity
                residual = positions[index] - predicted
                estimated_position = predicted + alphas * residual
                estimated_velocity += betas * residual / dt
                if index - start >= min(5, (stop - start) // 4):
                    error = estimated_velocity - feedback[index]
                    squared_errors += error * error
                    count += 1
        if count == 0:
            raise ValueError("資料不足，無法自動調適 α–β 參數。")
        return alphas, fractions, np.sqrt(squared_errors / count)

    alpha_min, alpha_max = 0.001, 1.999
    fraction_min, fraction_max = 0.001, 0.999
    alpha_grid = np.linspace(alpha_min, alpha_max, 25)
    fraction_grid = np.linspace(fraction_min, fraction_max, 25)
    alpha_step = (alpha_max - alpha_min) / 24
    fraction_step = (fraction_max - fraction_min) / 24
    best = None

    for refinement in range(4):
        alphas, fractions, errors = scores(alpha_grid, fraction_grid)
        best_index = int(np.argmin(errors))
        best = (float(alphas[best_index]), float(fractions[best_index]), float(errors[best_index]))
        if refinement < 3:
            alpha_step /= 4
            fraction_step /= 4
            alpha_grid = np.linspace(
                max(alpha_min, best[0] - alpha_step * 4),
                min(alpha_max, best[0] + alpha_step * 4),
                9,
            )
            fraction_grid = np.linspace(
                max(fraction_min, best[1] - fraction_step * 4),
                min(fraction_max, best[1] + fraction_step * 4),
                9,
            )

    alpha, fraction, rmse = best
    beta = fraction * (4 - 2 * alpha)
    return replace(settings, alpha=alpha, beta=beta), rmse


def true_runs(mask):
    return [(int(a), int(b)) for a, b in np.flatnonzero(np.diff(np.r_[False, mask, False])).reshape(-1, 2)]


def detect_motions(log, settings):
    movements = []
    for lo, hi in contiguous_blocks(len(log.time), log.gaps):
        values = log.speed_cmd[lo:hi]
        runs = sorted((lo + a, lo + b, direction)
                      for direction in (-1, 1)
                      for a, b in true_runs(direction * values > settings.motion_threshold))
        joined = []
        for a, b, sign in runs:
            if joined and joined[-1][2] == sign and a - joined[-1][1] <= settings.merge_gap:
                joined[-1] = (joined[-1][0], b, sign)
            else:
                joined.append((a, b, sign))
        sustained = [(a, b, sign) for a, b, sign in joined if b - a >= settings.min_motion]
        for i, (a, b, sign) in enumerate(sustained):
            move = Movement(len(movements) + 1, a, b, sign, lo, hi)
            move.floor = sustained[i - 1][1] if i else lo
            move.ceiling = sustained[i + 1][0] if i + 1 < len(sustained) else hi
            directed = sign * log.speed_cmd[a:b]
            q = len(directed) // 4
            center = directed[q:len(directed) - q]
            plateau = float(np.mean(center))
            central_is_flat = np.max(np.abs(center - plateau)) <= max(0.02 * plateau, 1.0)
            if not central_is_flat:
                # 錄製在平台中斷時，中央區域可能包含暫態。只在資料邊界採用穩定端點補認平台。
                edge = directed[:settings.min_motion] if a == lo else directed[-settings.min_motion:] if b == hi else None
                if edge is not None:
                    reference = float(np.mean(edge))
                    if np.max(np.abs(edge - reference)) <= max(0.01 * reference, 1.0):
                        plateau, central_is_flat = reference, True
            if plateau > 0 and central_is_flat:
                flat_runs = true_runs(np.abs(directed - plateau) <= max(0.01 * plateau, 1.0))
                if flat_runs:
                    left, right = max(flat_runs, key=lambda pair: pair[1] - pair[0])
                    if right - left >= settings.min_motion:
                        move.plateau, move.flat = plateau, (a + left, a + right)
            movements.append(move)
    return movements


def normalized_scurve(t, front, hold, back):
    """由 0 到 1 的分段速度，前後圓滑段為二次式，中段為一次式。"""
    duration = front + hold + back
    u = np.clip(t, 0.0, duration)
    scale = hold + (front + back) / 2
    rise = u**2 / (2 * front)
    middle = u - front / 2
    fall = scale - (duration - u)**2 / (2 * back)
    return np.where(u < front, rise, np.where(u < front + hold, middle, fall)) / scale


def cross_at(t, speed, level, rising, count, start, stop):
    if stop - start <= count:
        return None
    v = speed[start:stop]
    on_side = v >= level if rising else v <= level
    for a, b in true_runs(on_side):
        if a == 0:
            if b >= count:
                # 已在門檻的另一側開始記錄，無法內插出跨越時間。
                return None
            continue
        if b - a < count:
            continue
        j = start + a
        v0, v1 = float(speed[j - 1]), float(speed[j])
        time = t[j - 1] + (level - v0) * (t[j] - t[j - 1]) / (v1 - v0)
        return Crossing(j - 1, j, v0, v1, float(time))
    return None


def fit_ramp(time, desired, plateau, dt, start, stop, anchor, rough_ms):
    # 參數用樣本寬度表示，讓解算對不同取樣週期保持相近的數值尺度。
    relative = (time[start:stop] - time[anchor]) / dt
    y = desired / plateau
    if len(y) < 10 or np.sum(y <= 0.01) < 3 or np.sum(y >= 0.99) < 3:
        return None, "斜坡兩側有效基準不足"
    roughly = max(rough_ms / dt, 6.0)
    guess_front = max(0.75, 0.20 * roughly)
    initial = [-min(roughly * 0.10, -relative[0] * 0.80), guess_front,
               max(1, 0.60 * roughly), guess_front]
    lo = [relative[0], 0.25, 0.0, 0.25]
    hi = [min(relative[-1] - 2, max(2, 0.5 * roughly)), len(y), len(y), len(y)]
    if hi[0] <= lo[0] + 1e-6:
        return None, "斜坡候選點太靠近記錄邊界"
    solved = least_squares(
        lambda p: normalized_scurve(relative - p[0], *p[1:]) - y,
        np.clip(initial, np.asarray(lo) + 1e-7, np.asarray(hi) - 1e-7),
        bounds=(lo, hi), max_nfev=250, ftol=1e-10, xtol=1e-10, gtol=1e-10,
    )
    origin, front, hold, back = solved.x
    error = normalized_scurve(relative - origin, front, hold, back) - y
    rmse = float(np.sqrt(np.mean(error**2)) * plateau)
    if not solved.success or rmse > max(1.0, 0.02 * plateau) or np.max(np.abs(error)) * plateau > max(2.0, 0.08 * plateau):
        return None, "S 曲線擬合誤差較大"
    if min(front, back) < 1.0:
        return None, "圓滑子段短於 1 筆，無法辨識 S 曲線"
    if origin - relative[0] < 2 or relative[-1] - (origin + front + hold + back) < 2:
        return None, "擬合接點距記錄邊界太近"
    fit = Fit(float(time[anchor] + origin * dt), float(front * dt), float(hold * dt), float(back * dt),
              rmse, float(plateau / ((hold + (front + back) / 2) * dt)))
    return fit, "已擬合"


def analyze_ramps(log, estimate, moves, settings):
    ramps = []
    dt = settings.tick_ms
    signal_sets = {1: (log.speed_cmd, log.speed_fb, estimate.velocity),
                   -1: (-log.speed_cmd, -log.speed_fb, -estimate.velocity)}
    absolute_feedback = np.abs(log.speed_fb)
    for move in moves:
        for rising in (True, False):
            fractions = (.05, .50, .95) if rising else (.95, .50, .05)
            anchor = move.first if rising else move.stop - 1
            empty = [[None] * 3 for _ in SIGNALS]
            if move.flat is None:
                ramps.append(Ramp(move, rising, anchor, None, "平台速率無法可靠辨識", fractions, empty))
                continue

            flat_start, flat_stop = move.flat
            mid = (flat_start + flat_stop) // 2
            if not rising:
                anchor = flat_stop
                if anchor >= move.ceiling:
                    ramps.append(Ramp(move, rising, move.stop - 1, None, "記錄在平台內結束", fractions, empty))
                    continue
            rough = max(6, flat_start - move.first if rising else move.stop - flat_stop)
            left = max(move.floor if rising else mid, anchor - max(25, math.ceil(rough * .4)))
            right = min(mid + 1 if rising else move.ceiling,
                        anchor + max(220, math.ceil(rough * 1.6) + 30))
            measures = signal_sets[move.direction]
            directed = measures[0][left:right]
            desired = directed if rising else move.plateau - directed
            fitted, state = fit_ramp(log.time, desired, move.plateau, dt, left, right, anchor, rough * dt)
            measure_stop = min(move.ceiling, anchor + max(450, rough * 3 + 50))
            if rising:
                measure_stop = min(measure_stop, flat_stop)
            hits = [[cross_at(log.time, signal, move.plateau * f, rising,
                              1 if i == 0 else settings.crossing_samples, left, measure_stop)
                     for f in fractions] for i, signal in enumerate(measures)]
            near = None if rising else cross_at(log.time, absolute_feedback, move.plateau * .01,
                                                False, 20, left, measure_stop)
            peak_stop = min(move.ceiling, flat_stop if rising else move.ceiling,
                            anchor + max(250, rough * 2 + 40))
            peak_start = max(left, anchor - max(18, rough // 5))
            span = move.direction * estimate.acceleration[peak_start:peak_stop]
            valid = span[np.isfinite(span)]
            peak = float((np.max if rising else np.min)(valid)) if valid.size else None
            ramps.append(Ramp(move, rising, anchor, fitted, state, fractions, hits, near, peak))
    return ramps


def width(crossings):
    return crossings[2].time - crossings[0].time if crossings[0] and crossings[2] else None


def lag(observed, command):
    return observed[1].time - command[1].time if observed[1] and command[1] else None


def finite(value):
    if isinstance(value, (float, np.floating)) and not math.isfinite(value):
        return None
    return value.item() if isinstance(value, np.generic) else value


def write_row(sheet, row, values):
    sheet.write_row(row, 0, [finite(value) for value in values])


def write_explanations(sheet, content, formats):
    sheet.hide_gridlines(2)
    sheet.set_column("A:A", 26)
    sheet.set_column("B:B", 112)
    sheet.freeze_panes(1, 0)
    sheet.write_row(0, 0, ["項目", "定義與判讀"], formats["header"])
    for i, (label, description) in enumerate(content, 1):
        sheet.write(i, 0, label, formats["label"])
        sheet.write(i, 1, description, formats["wrap"])
        sheet.set_row(i, 45 if len(description) > 75 else 31)


def prepare_table(sheet, headers, formats, widths=None):
    sheet.hide_gridlines(2)
    sheet.freeze_panes(1, 1)
    sheet.set_row(0, 35)
    for col in range(len(headers)):
        width = widths[col] if widths and col < len(widths) else 17
        sheet.set_column(col, col, width, formats["number"])
    sheet.write_row(0, 0, headers, formats["header"])


def export_detail(book, log, estimate, moves, ramps, formats, kalman_estimate=None):
    headers = [
        "來源 Excel 列", "原 index", "tick", "相對時間\nms", "Δt\nms",
        "Xcmd\npulse", "Xfb\npulse", "原 Vxcmd", "原 Vxfb", "原 Following_error",
        "分析 Vxcmd\npulse/ms", "分析 Vxfb\npulse/ms", "分析追隨誤差\npulse",
        "預測位置\npulse", "位置殘差\npulse", "α–β 位置\npulse", "α–β 速度\npulse/ms",
        "α–β 加速度\npulse/ms²", "命令加速度\npulse/ms²", "運動段次", "S 曲線階段",
    ]
    if kalman_estimate is not None:
        headers.extend((
            "卡曼估測位置\npulse", "卡曼估測速度\npulse/ms", "卡曼估測加速度\npulse/ms²",
        ))
        if kalman_estimate.adaptive:
            headers.extend(("卡曼 Q\n程序雜訊強度", "卡曼 R\n位置量測變異數"))
    length = len(log.time)
    segment_ids = np.zeros(length, dtype=np.int32)
    stages = np.zeros(length, dtype=np.int8)
    stage_names = ("", "加速度建立", "定加速度", "加速度回零", "等速",
                   "減速度建立", "定減速度", "減速度回零")
    for i, move in enumerate(moves):
        acc, dec = ramps[2 * i:2 * i + 2]
        start, stop = move.first, move.stop
        if acc.fit:
            start = int(np.searchsorted(log.time, acc.fit.start))
        if dec.fit:
            stop = int(np.searchsorted(log.time, dec.fit.boundaries[3], side="right"))
        segment_ids[start:stop] = move.number
        for ramp, stage_start in ((acc, 1), (dec, 5)):
            if ramp.fit:
                bounds = np.searchsorted(log.time, ramp.fit.boundaries)
                for j, (first, last) in enumerate(zip(bounds[:-1], bounds[1:])):
                    stages[first:last] = stage_start + j
        if acc.fit and dec.fit:
            a = np.searchsorted(log.time, acc.fit.boundaries[3])
            b = np.searchsorted(log.time, dec.fit.start)
            stages[a:b] = 4

    page_size = EXCEL_ROWS - 1
    for part, start in enumerate(range(0, length, page_size), 1):
        stop = min(length, start + page_size)
        sheet = book.add_worksheet("逐筆估算" if part == 1 else f"逐筆估算 {part}")
        prepare_table(sheet, headers, formats)
        for k in range(start, stop):
            idx, xcmd, xfb, vcmd, vfb, following, tick = log.data[k]
            values = (
                log.source_rows[k], idx, tick, log.time[k], log.interval[k] if k else None,
                xcmd, xfb, vcmd, vfb, following, log.speed_cmd[k], log.speed_fb[k], log.following[k],
                estimate.predicted[k], estimate.residual[k], estimate.position[k], estimate.velocity[k],
                estimate.acceleration[k], estimate.command_acceleration[k],
                int(segment_ids[k]) or None, stage_names[stages[k]] or None,
            )
            if kalman_estimate is not None:
                values += (
                    kalman_estimate.position[k],
                    kalman_estimate.velocity[k],
                    kalman_estimate.acceleration[k],
                )
                if kalman_estimate.adaptive:
                    values += (
                        kalman_estimate.process_noise_values[k],
                        kalman_estimate.measurement_variance_values[k],
                    )
            write_row(sheet, k - start + 1, values)
        sheet.autofilter(0, 0, stop - start, len(headers) - 1)


def export_summary(sheet, log, moves, ramps, formats):
    headers = [
        "運動段次", "編碼器方向", "平台速率\npulse/ms", "起始樣本\nms",
        "末端樣本\nms", "加速全斜坡\nms", "減速全斜坡\nms", "等速時間\nms",
        "加速回授延遲\n50% ms", "減速回授延遲\n50% ms",
        "峰值追隨誤差\n|pulse|", "判讀",
    ]
    prepare_table(sheet, headers, formats, [12, 13, 18, 19, 19, 19, 19, 17, 20, 20, 20, 66])
    for row, move in enumerate(moves, 1):
        acc, dec = ramps[2 * (row - 1):2 * row]
        flat = (dec.fit.start - acc.fit.boundaries[3]) if acc.fit and dec.fit else None
        status = f"加速 {acc.status}，減速 {dec.status}"
        if flat is not None and flat < 0:
            flat = None
            status = "兩個擬合斜坡重疊，等速時間無法判定"
        peak_error = np.max(np.abs(log.following[move.first:move.stop]))
        write_row(sheet, row, [
            move.number, "正向" if move.direction > 0 else "負向", move.plateau,
            log.time[move.first], log.time[move.stop - 1],
            acc.fit.boundaries[3] - acc.fit.start if acc.fit else None,
            dec.fit.boundaries[3] - dec.fit.start if dec.fit else None,
            flat, lag(acc.crossings[1], acc.crossings[0]), lag(dec.crossings[1], dec.crossings[0]),
            peak_error, status,
        ])
    if not moves:
        sheet.write(1, 0, "未偵測到符合目前門檻的連續命令運動。", formats["label"])
    else:
        sheet.autofilter(0, 0, len(moves), len(headers) - 1)


def export_ramps(sheet, thresholds, log, ramps, formats):
    headers = [
        "段次", "斜坡", "辨識狀態", "方向", "平台速率\npulse/ms",
        "t0 起點\nms", "t1 接點\nms", "t2 接點\nms", "t3 終點\nms",
        "建立\nms", "維持\nms", "回零\nms", "完整斜坡\nms",
        "命令 5%–95%\nms", "原始回授 5%–95%\nms", "α–β 5%–95%\nms",
        "原始回授延遲\n50% ms", "α–β 回授延遲\n50% ms", "α–β 額外延遲\n50% ms",
        "擬合 RMSE\npulse/ms", "命令定加速度\n|pulse/ms²|", "回授峰值加速度\npulse/ms²",
        "原始回授近零速\nms", "命令結束至近零\nms",
    ]
    prepare_table(sheet, headers, formats, [10, 10, 37, 10])
    threshold_headers = [
        "段次", "斜坡", "訊號", "門檻\n% 平台", "門檻速率\npulse/ms",
        "前一筆 Excel 列", "跨越 Excel 列", "前一筆時間\nms", "跨越時間\nms",
        "前一筆速率\npulse/ms", "跨越速率\npulse/ms", "門檻內插時間\nms",
    ]
    prepare_table(thresholds, threshold_headers, formats, [10, 10, 21])
    threshold_row = 1
    for row, ramp in enumerate(ramps, 1):
        fit = ramp.fit
        command, raw_fb, filtered = ramp.crossings
        near = ramp.near_zero.time if ramp.near_zero else None
        width_fields = [width(hit) for hit in ramp.crossings]
        raw_lag, filtered_lag = lag(raw_fb, command), lag(filtered, command)
        bounds = list(fit.boundaries) if fit else [None] * 4
        durations = [fit.rise, fit.hold, fit.fall, bounds[3] - bounds[0]] if fit else [None] * 4
        write_row(sheet, row, [
            ramp.move.number, ramp.label, ramp.status, ramp.move.direction, ramp.move.plateau,
            *bounds, *durations, *width_fields, raw_lag, filtered_lag,
            filtered_lag - raw_lag if filtered_lag is not None and raw_lag is not None else None,
            fit.rmse if fit else None, fit.peak if fit else None, ramp.peak_feedback, near,
            near - bounds[3] if near is not None and fit else None,
        ])
        items = [(signal, fraction, hit) for signal, hits in zip(SIGNALS, ramp.crossings)
                 for fraction, hit in zip(ramp.fractions, hits)]
        if not ramp.accelerating:
            items.append(("|Vxfb| 近零速", .01, ramp.near_zero))
        for signal, fraction, hit in items:
            level = ramp.move.plateau * fraction if ramp.move.plateau is not None else None
            evidence = ([log.source_rows[hit.before], log.source_rows[hit.after],
                         log.time[hit.before], log.time[hit.after], hit.speed_before, hit.speed_after, hit.time]
                        if hit else [None] * 7)
            write_row(thresholds, threshold_row, [ramp.move.number, ramp.label, signal, fraction * 100, level, *evidence])
            threshold_row += 1
    if ramps:
        sheet.autofilter(0, 0, len(ramps), len(headers) - 1)
        thresholds.autofilter(0, 0, threshold_row - 1, len(threshold_headers) - 1)


def sample_range(start, stop, limit):
    if stop <= start:
        return np.array([], dtype=int)
    if stop - start <= limit:
        return np.arange(start, stop)
    return np.linspace(start, stop - 1, limit, dtype=int)


def add_chart(book, target, source, start, values, columns, title, axis_x, axis_y, cell):
    chart = book.add_chart({"type": "scatter", "subtype": "straight"})
    colors = {
        1: "#2468A0", 2: "#845A9E", 3: "#087E78", 4: "#2468A0",
        5: "#087E78", 6: "#D17A00", 7: "#D17A00",
    }
    end = start + len(values)
    categories = [finite(value) for value in values[:, 0]]
    for col in columns:
        chart.add_series({
            "name": [source.name, start, col],
            "categories": [source.name, start + 1, 0, end, 0],
            "values": [source.name, start + 1, col, end, col],
            "categories_data": categories,
            "values_data": [finite(value) for value in values[:, col]],
            "line": {"color": colors[col], "width": 1.5, "dash_type": "dash" if col == 2 else "solid"},
            "marker": {"type": "none"},
        })
    chart.set_title({"name": title})
    chart.set_x_axis({"name": axis_x, "num_format": "0.0"})
    chart.set_y_axis({"name": axis_y, "num_format": "0.0"})
    chart.set_legend({"position": "top"})
    chart.set_size({"width": 590, "height": 296})
    target.insert_chart(cell, chart)


def export_charts(book, charts, source, log, estimate, ramps, settings, formats,
                  kalman_estimate=None):
    charts.hide_gridlines(2)
    charts.write("A1", "全程速度與各斜坡局部曲線", formats["title"])
    description = "逐筆估算保留全部樣本。圖表數據頁保留抽樣點，局部圖按運動方向正規化。"
    if kalman_estimate is not None:
        description += "全程與局部圖包含卡曼速度及加速度。"
    charts.write("A2", description)
    charts.set_column("A:T", 13)
    header = ["相對時間 ms", "命令速度", "原始回授", "α–β 速度",
              "命令加速度", "α–β 加速度"]
    if kalman_estimate is not None:
        header.extend(("卡曼速度", "卡曼加速度"))
    prepare_table(source, header, formats, [20] * len(header))
    overview_samples = sample_range(0, len(log.time), 2500)
    overview = np.column_stack((log.time[overview_samples], log.speed_cmd[overview_samples],
                                log.speed_fb[overview_samples], estimate.velocity[overview_samples]))
    if kalman_estimate is not None:
        overview = np.column_stack((overview, kalman_estimate.velocity[overview_samples]))
    for i, values in enumerate(overview, 1):
        write_row(source, i, values)
    speed_columns = (1, 2, 3, 4) if kalman_estimate is not None else (1, 2, 3)
    add_chart(book, charts, source, 0, overview, speed_columns,
              "全程 X 軸速度", "相對時間（ms）", "速度（pulse/ms）", "A4")

    row = len(overview_samples) + 2
    displayed = [r for r in ramps if r.fit or any(r.crossings[0])][:settings.chart_ramps]
    for i, ramp in enumerate(displayed):
        fit = ramp.fit
        if fit:
            t0, _, _, t3 = fit.boundaries
            duration = t3 - t0
        else:
            t0, duration = log.time[ramp.anchor], 100 * settings.tick_ms
            t3 = t0 + duration
        start = max(ramp.move.block_start, np.searchsorted(log.time, t0 - max(24 * settings.tick_ms, duration * .20)))
        stop = min(ramp.move.block_stop, np.searchsorted(log.time, t3 + max(90 * settings.tick_ms, duration * .6), side="right"))
        samples = sample_range(start, stop, 360)
        if len(samples) < 2:
            continue
        source.write_row(row, 0, header, formats["header"])
        sign = ramp.move.direction
        local = np.column_stack((log.time[samples] - t0, sign * log.speed_cmd[samples],
                                 sign * log.speed_fb[samples], sign * estimate.velocity[samples],
                                 sign * estimate.command_acceleration[samples], sign * estimate.acceleration[samples]))
        if kalman_estimate is not None:
            local = np.column_stack((
                local, sign * kalman_estimate.velocity[samples],
                sign * kalman_estimate.acceleration[samples],
            ))
        for j, values in enumerate(local, row + 1):
            write_row(source, j, values)
        chart_row = 21 + i * 17
        prefix = f"第 {ramp.move.number} 段{ramp.label}"
        x_label = "相對擬合起點（ms）" if fit else "相對候選點（ms）"
        local_speed_columns = (1, 2, 3, 6) if kalman_estimate is not None else (1, 2, 3)
        local_acceleration_columns = (4, 5, 7) if kalman_estimate is not None else (4, 5)
        add_chart(book, charts, source, row, local, local_speed_columns, prefix + "速度",
                  x_label, "速率（pulse/ms）", f"A{chart_row}")
        add_chart(book, charts, source, row, local, local_acceleration_columns, prefix + "加速度",
                  x_label, "加速度（pulse/ms²）", f"I{chart_row}")
        row += len(samples) + 2

    if kalman_estimate is not None:
        acceleration_row = row
        acceleration_header = ("相對時間 ms", "命令加速度", "α–β 加速度", "卡曼加速度")
        source.write_row(acceleration_row, 0, acceleration_header, formats["header"])
        acceleration_values = np.column_stack((
            log.time[overview_samples],
            estimate.command_acceleration[overview_samples],
            estimate.acceleration[overview_samples],
            kalman_estimate.acceleration[overview_samples],
        ))
        for data_row, values in enumerate(acceleration_values, acceleration_row + 1):
            write_row(source, data_row, values)
        add_chart(book, charts, source, acceleration_row, acceleration_values,
                  (1, 2, 3), "全程 X 軸加速度", "相對時間（ms）",
                  "加速度（pulse/ms²）", "I4")


def export_notes(sheet, source, log, estimate, settings, formats, adaptive_rmse=None,
                 kalman_estimate=None):
    raw = log.data
    index_gaps = int(np.count_nonzero(np.diff(raw[:, 0]) != 1))
    cmd_diff = int(np.count_nonzero(~np.isclose(np.diff(raw[:, 1]) / log.interval[1:], log.speed_cmd[1:])))
    fb_diff = int(np.count_nonzero(~np.isclose(np.diff(raw[:, 2]) / log.interval[1:], log.speed_fb[1:])))
    following_diff = int(np.count_nonzero(~np.isclose(raw[:, 1] - raw[:, 2], log.following)))
    rmse = np.sqrt(np.mean((estimate.velocity - log.speed_fb)**2))
    content = [
        ("來源", f"{source.name}，工作表「{log.sheet}」，共 {len(raw):,} 筆。資料表保留原 Excel 列號。"),
        ("時間與單位", f"tick 乘 {settings.tick_ms:g} ms/tick，以首筆為零。原速度單位設定為 {settings.velocity_unit}，分析速度以 pulse/ms 表示，加速度以 pulse/ms² 表示。"),
        ("時間完整性", f"index 不連續 {index_gaps:,} 處，tick 不連續 {len(log.gaps):,} 處。每個 tick 連續區塊獨立初始化，濾波、差分與分段皆在區塊內執行。"),
        ("資料核對", f"分析 Vxcmd 與位置差商有 {cmd_diff:,} 筆差異，Vxfb 有 {fb_diff:,} 筆差異，追隨誤差有 {following_diff:,} 筆差異。數量含可比較的樣本。"),
        ("缺值處理", f"以位置差商補入 Vxcmd {log.filled[0]:,} 筆、Vxfb {log.filled[1]:,} 筆。以 Xcmd−Xfb 補入追隨誤差 {log.filled[2]:,} 筆。原值欄保留空白。"),
        ("初始化", "每個連續區塊首筆位置取 Xfb，速度取分析 Vxfb。預測位置及殘差從次筆開始。"),
        ("α–β 預測", "x⁻[k] = x̂[k−1] + Δt × v̂[k−1]。殘差 r[k] = Xfb[k] − x⁻[k]。"),
        ("α–β 更新", f"α = {settings.alpha:g}，β = {settings.beta:g}。x̂[k] = x⁻[k] + α × r[k]。v̂[k] = v̂[k−1] + β × r[k]/Δt。"),
        ("參數設定", (f"自動調適：最小化估計速度與 Vxfb 的 RMSE，調適樣本 RMSE 為 {adaptive_rmse:.6g} pulse/ms。"
                      if adaptive_rmse is not None else "手動設定 α、β。")),
        ("卡曼濾波", ("啟用三狀態常加速度模型，狀態為位置、速度、加速度；以 Xfb 為位置觀測，"
                      f"基準 Q（白雜訊 jerk 程序雜訊強度）={kalman_estimate.process_noise:.8g}，"
                      f"基準 R（位置量測雜訊變異數）={kalman_estimate.measurement_variance:.8g}。"
                      + (f"自適應模式：IAE 以 0.05 更新率估測 R，限制在基準 R 的 0.1 至 10 倍；"
                         f"依 |Vxcmd| 是否超過 {settings.motion_threshold:g} pulse/ms 判斷運動，"
                         "運動時 Q=基準 Q，靜止時 Q=基準 Q×0.1。Q、R 不以同一誤差訊號同時自由最佳化；"
                         "逐筆採用值列於逐筆估算工作表。"
                         if kalman_estimate.adaptive else "")
                      + "每個連續 tick 區塊獨立初始化，估測值列於逐筆估算工作表及速度與加速度圖。"
                      if kalman_estimate is not None else "未啟用。")),
        ("離線加速度", f"使用前後各 {settings.half_window} 筆中心差分，實際跨度為 {settings.half_window * 2 * settings.tick_ms:g} ms。區塊首尾缺少足夠資料的位置留空。"),
        ("速度差異", f"全程 α–β 速度與分析 Vxfb 的 RMSE 為 {rmse:.3f} pulse/ms，包含平滑與暫態延遲。"),
        ("運動識別", f"|Vxcmd| 超過 {settings.motion_threshold:g} pulse/ms 持續至少 {settings.min_motion} 筆。合併同向且相隔不超過 {settings.merge_gap} 筆的命令。"),
        ("平台基準", "取運動中間約 50% 的平均速率。窗口須落在均值 2% 或 1 pulse/ms 內。若記錄在運動中截斷，另檢查穩定端點是否可作平台基準。"),
        ("命令 S 曲線", "最小平方法擬合兩端的二次速度曲線與中間的一次速度曲線。t0 到 t3 分別依序標示起點、定斜率起點、回零起點、全斜坡終點。"),
        ("子段意義", "t1−t0 是加速／減速度建立，t2−t1 是定加速度或定減速度，t3−t2 是加速度回零。等速時間從加速 t3 到減速 t0。"),
        ("適用條件", "完整時間依賴穩定平台、靜止基準和足夠樣本。三角形速度、連續變速或記錄截斷時，無法辨識的欄位留空，可查看辨識狀態與可用門檻。"),
        ("速度門檻", f"加速取 5%→95%，減速取 95%→5%。命令取第一次跨越，回授與 α–β 需連續 {settings.crossing_samples} 筆。時間用相鄰樣本線性內插。"),
        ("50% 延遲", "同一命令平台的 50% 為基準。回授跨越時間減命令跨越時間。α–β 的額外延遲為該估計與原始回授的跨越時間差。"),
        ("近零速", "原始回授的絕對速度須進入該段平台速率的 1% 以內並持續 20 筆。此時間點可與命令減速結束 t3 相減。"),
        ("精度與圖表", "小數時間起於模型擬合與內插。全程圖與局部圖採等間距抽樣，逐筆估算頁包含完整資料。未附機構導程時使用 pulse 為位置單位。"),
    ]
    write_explanations(sheet, content, formats)


def write_nc_code(sheet, nc_code, formats):
    lines = nc_code.splitlines()
    if len(lines) > EXCEL_ROWS - 1:
        raise ValueError(f"NC code 超過 Excel 單一工作表可容納的 {EXCEL_ROWS - 1:,} 行。")
    for line_number, line in enumerate(lines, 1):
        if len(line) > 32767:
            raise ValueError(f"NC code 第 {line_number} 行超過 Excel 單一儲存格的 32,767 字元限制。")
    sheet.write_row(0, 0, ("行號", "NC Code"), formats["header"])
    sheet.set_column("A:A", 10)
    sheet.set_column("B:B", 110)
    sheet.freeze_panes(1, 0)
    if not lines:
        sheet.write_string(1, 1, "未提供 NC code。")
        return
    for row, line in enumerate(lines, 1):
        sheet.write_number(row, 0, row)
        sheet.write_string(row, 1, line)


def write_report(path, source, log, estimate, moves, ramps, settings, nc_code="",
                 adaptive_rmse=None, kalman_estimate=None):
    # 同一目錄的暫存檔讓最後一步可原子替換，例外時保留既有報告與來源。
    with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".axis_scurve_", suffix=".xlsx", delete=False) as temporary:
        pending = Path(temporary.name)
    try:
        with xlsxwriter.Workbook(pending, {"constant_memory": True, "strings_to_formulas": False,
                                         "strings_to_urls": False}) as book:
            book.set_properties({"title": "X 軸 α–β 濾波與 S 曲線分析"})
            formats = {
                "header": book.add_format({"bold": True, "bg_color": "#163450", "font_color": "white",
                                           "valign": "vcenter", "align": "center", "text_wrap": True}),
                "number": book.add_format({"num_format": "0.000;-0.000;0"}),
                "label": book.add_format({"bold": True, "font_color": "#163450", "valign": "vcenter"}),
                "wrap": book.add_format({"text_wrap": True, "valign": "vcenter"}),
                "title": book.add_format({"bold": True, "font_color": "#163450", "font_size": 16}),
            }
            summary = book.add_worksheet("分析總覽")
            s_curves = book.add_worksheet("S曲線分段")
            crossings = book.add_worksheet("門檻明細")
            charts = book.add_worksheet("速度與加速度圖")
            notes = book.add_worksheet("方法與設定")
            graph_data = book.add_worksheet("圖表數據")
            nc_sheet = book.add_worksheet("NC Code")
            export_summary(summary, log, moves, ramps, formats)
            export_ramps(s_curves, crossings, log, ramps, formats)
            export_charts(book, charts, graph_data, log, estimate, ramps, settings, formats,
                          kalman_estimate)
            export_notes(notes, source, log, estimate, settings, formats,
                         adaptive_rmse, kalman_estimate)
            export_detail(book, log, estimate, moves, ramps, formats, kalman_estimate)
            write_nc_code(nc_sheet, nc_code, formats)
        try:
            pending.replace(path)
        except PermissionError as exc:
            raise PermissionError(
                f"分析與報告內容已完成，但無法寫入目標檔案：{path}\n"
                "若該報告正由 Excel 開啟，請先關閉後重新分析；"
                "也請確認輸出資料夾可寫入，或使用 -o 指定其他檔名。"
            ) from exc
    finally:
        pending.unlink(missing_ok=True)


def args_parser():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("input", nargs="?", type=Path,
                        help="來源 .xlsx 或 .xlsm。省略時開啟選檔視窗，或在終端機輸入路徑")
    parser.add_argument("--no-dialog", action="store_true", help="省略來源檔時直接在終端機輸入路徑")
    parser.add_argument("-o", "--output", type=Path, help="輸出 .xlsx。預設在來源旁建立 _scurve_analysis.xlsx")
    parser.add_argument("--sheet", help="資料工作表名稱。預設搜尋包含必要欄位的工作表")
    parser.add_argument("--tick-ms", type=float, default=1.0, help="每 tick 的毫秒數，預設 1")
    parser.add_argument("--velocity-unit", choices=["pulse/ms", "pulse/sample"], default="pulse/sample",
                        help="原始速度欄單位，預設脈衝差 pulse/sample。1 ms 取樣時兩者數值相同")
    parser.add_argument("--alpha", type=float, default=.64, help="位置更新增益，預設 0.64")
    parser.add_argument("--beta", type=float, default=.16, help="速度更新增益，預設 0.16")
    parser.add_argument("--auto-alpha-beta", action="store_true", help="依估計速度與 Vxfb 的 RMSE 自動調適 α、β")
    parser.add_argument("--kalman-filter", action="store_true", help="執行卡曼濾波並輸出估測位置、速度與加速度")
    parser.add_argument("--kalman-q", type=float, help="卡曼程序雜訊強度 Q；省略時依資料估算")
    parser.add_argument("--kalman-r", type=float, help="卡曼位置量測雜訊變異數 R；省略時依資料估算")
    parser.add_argument("--adaptive-kalman", action="store_true",
                        help="自適應卡曼：IAE 調整 R，並依命令速度運動狀態調整 Q")
    parser.add_argument("--nc-code-file", type=Path, help="將 UTF-8 NC code 文字檔加入輸出活頁簿")
    parser.add_argument("--half-window", type=int, default=3, help="加速度中心差分前後各取幾筆，預設 3")
    parser.add_argument("--motion-threshold", type=float, default=5, help="運動偵測的命令速率門檻 pulse/ms，預設 5")
    parser.add_argument("--min-motion", type=int, default=8, help="最短運動與平台樣本數，預設 8")
    parser.add_argument("--merge-gap", type=int, default=2, help="同向運動可合併的間隔筆數，預設 2")
    parser.add_argument("--crossing-samples", type=int, default=5, help="回授通過門檻後的連續確認筆數，預設 5")
    parser.add_argument("--chart-ramps", type=int, default=12, help="局部速度與加速度圖最多畫幾個斜坡，預設 12")
    parser.add_argument("--overwrite", action="store_true", help="允許覆寫既有輸出報告")
    return parser


def select_windows_file():
    """透過 Windows 內建的 WinForms 開啟選檔視窗，避免依賴 Tcl/Tk。"""
    script = """
        $ErrorActionPreference = 'Stop'
        if (-not [Environment]::UserInteractive) { exit 1 }
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        try {
            $dialog.Title = '選取 X 軸量測 Excel'
            $dialog.Filter = 'Excel 活頁簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm'
            $dialog.InitialDirectory = [Environment]::CurrentDirectory
            $dialog.CheckFileExists = $true
            $dialog.Multiselect = $false
            if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { exit 2 }
            [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
            [Console]::WriteLine($dialog.FileName)
        } finally {
            $dialog.Dispose()
        }
    """
    powershell = (Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
                  / "WindowsPowerShell" / "v1.0" / "powershell.exe")
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    finished = subprocess.run(
        [str(powershell), "-NoLogo", "-NoProfile", "-STA", "-EncodedCommand", encoded],
        stdin=subprocess.DEVNULL, capture_output=True,
        creationflags=subprocess.CREATE_NO_WINDOW, check=False,
    )
    if finished.returncode == 2:
        return None
    if finished.returncode != 0:
        raise RuntimeError("Windows 選檔視窗啟動失敗")
    filename = finished.stdout.decode("utf-8-sig").rstrip("\r\n")
    if not filename:
        raise RuntimeError("Windows 選檔視窗未傳回檔案路徑")
    return Path(filename)


def select_tk_file():
    import tkinter as tk
    from tkinter import filedialog

    try:
        root = tk.Tk()
        root.withdraw()
        try:
            filename = filedialog.askopenfilename(parent=root, title="選取 X 軸量測 Excel",
                                                 filetypes=[("Excel 活頁簿", "*.xlsx *.xlsm")])
        finally:
            root.destroy()
    except tk.TclError as exc:
        raise RuntimeError("選檔視窗啟動失敗") from exc
    return Path(filename) if filename else None


def select_input(no_dialog):
    if not no_dialog:
        picker = select_windows_file if sys.platform == "win32" else select_tk_file
        try:
            return picker()
        except (ImportError, OSError, RuntimeError, UnicodeError):
            print("無法開啟選檔視窗，請改用下方提示輸入路徑。", file=sys.stderr, flush=True)
    try:
        filename = input("請貼上 Excel 檔案的完整路徑，或直接按 Enter 取消：").strip()
    except EOFError:
        raise ValueError("目前無法讀取終端機輸入，請在執行指令的末尾加上 Excel 檔案路徑。") from None
    except KeyboardInterrupt:
        return None
    if len(filename) >= 2 and filename[0] == filename[-1] and filename[0] in {"'", '"'}:
        filename = filename[1:-1]
    return Path(filename) if filename else None


def launch_windows_gui():
    script = r"""
        $ErrorActionPreference = 'Stop'
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        function Quote-Argument([string]$value) {
            $builder = New-Object System.Text.StringBuilder
            [void]$builder.Append('"')
            $slashes = 0
            foreach ($character in $value.ToCharArray()) {
                if ($character -eq '\') {
                    $slashes++
                } elseif ($character -eq '"') {
                    [void]$builder.Append([string]::new([char]92, 2 * $slashes + 1))
                    [void]$builder.Append('"')
                    $slashes = 0
                } else {
                    if ($slashes -gt 0) { [void]$builder.Append([string]::new([char]92, $slashes)) }
                    [void]$builder.Append($character)
                    $slashes = 0
                }
            }
            if ($slashes -gt 0) { [void]$builder.Append([string]::new([char]92, 2 * $slashes)) }
            [void]$builder.Append('"')
            return $builder.ToString()
        }

        function Set-ControlFontSize(
            [System.Windows.Forms.Control]$control,
            [single]$size
        ) {
            $control.Font = [System.Drawing.Font]::new(
                $control.Font.FontFamily, $size, $control.Font.Style)
            foreach ($child in $control.Controls) {
                Set-ControlFontSize $child $size
            }
        }

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'X 軸 α–β 濾波與 S 曲線分析'
        $form.Size = New-Object System.Drawing.Size(740, 760)
        $form.MinimumSize = New-Object System.Drawing.Size(680, 700)
        $form.StartPosition = 'CenterScreen'
        $form.Font = [System.Drawing.Font]::new($form.Font.FontFamily, 12)

        $tabs = New-Object System.Windows.Forms.TabControl
        $tabs.Dock = 'Fill'
        $basic = New-Object System.Windows.Forms.TabPage
        $basic.Text = '基本功能區'
        $about = New-Object System.Windows.Forms.TabPage
        $about.Text = '軟體功能說明'
        [void]$tabs.TabPages.Add($basic)
        [void]$tabs.TabPages.Add($about)
        [void]$form.Controls.Add($tabs)

        $fontLabel = New-Object System.Windows.Forms.Label
        $fontLabel.Text = '字體大小（空白使用 12）'
        $fontLabel.Location = New-Object System.Drawing.Point(24, 18)
        $fontLabel.AutoSize = $true
        [void]$basic.Controls.Add($fontLabel)
        $fontBox = New-Object System.Windows.Forms.TextBox
        $fontBox.Location = New-Object System.Drawing.Point(210, 14)
        $fontBox.Width = 80
        [void]$basic.Controls.Add($fontBox)
        $fontButton = New-Object System.Windows.Forms.Button
        $fontButton.Text = '變更'
        $fontButton.Location = New-Object System.Drawing.Point(305, 12)
        $fontButton.AutoSize = $true
        [void]$basic.Controls.Add($fontButton)
        $fontButton.add_Click({
            $size = 12
            $text = $fontBox.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($text) -and
                (-not [int]::TryParse($text, [ref]$size) -or $size -lt 6 -or $size -gt 72)) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    '字體大小請輸入 6 到 72 的整數；留空則使用 12。',
                    '字體大小無效', [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if ([string]::IsNullOrWhiteSpace($text)) { $size = 12 }
            Set-ControlFontSize $form $size
        })

        $alphaLabel = New-Object System.Windows.Forms.Label
        $alphaLabel.Text = 'α（位置更新增益）'
        $alphaLabel.Location = New-Object System.Drawing.Point(24, 58)
        $alphaLabel.AutoSize = $true
        [void]$basic.Controls.Add($alphaLabel)
        $alphaBox = New-Object System.Windows.Forms.TextBox
        $alphaBox.Text = '0.64'
        $alphaBox.Location = New-Object System.Drawing.Point(190, 55)
        $alphaBox.Width = 150
        [void]$basic.Controls.Add($alphaBox)

        $betaLabel = New-Object System.Windows.Forms.Label
        $betaLabel.Text = 'β（速度更新增益）'
        $betaLabel.Location = New-Object System.Drawing.Point(24, 98)
        $betaLabel.AutoSize = $true
        [void]$basic.Controls.Add($betaLabel)
        $betaBox = New-Object System.Windows.Forms.TextBox
        $betaBox.Text = '0.16'
        $betaBox.Location = New-Object System.Drawing.Point(190, 95)
        $betaBox.Width = 150
        [void]$basic.Controls.Add($betaBox)

        $autoCheck = New-Object System.Windows.Forms.CheckBox
        $autoCheck.Text = '自調適 α–β（以 Vxfb RMSE 最小化）'
        $autoCheck.Location = New-Object System.Drawing.Point(24, 135)
        $autoCheck.AutoSize = $true
        [void]$basic.Controls.Add($autoCheck)
        $autoCheck.add_CheckedChanged({
            $alphaBox.Enabled = -not $autoCheck.Checked
            $betaBox.Enabled = -not $autoCheck.Checked
        })

        $kalmanCheck = New-Object System.Windows.Forms.CheckBox
        $kalmanCheck.Text = '啟用卡曼濾波並輸出估測位置、速度與加速度'
        $kalmanCheck.Location = New-Object System.Drawing.Point(24, 162)
        $kalmanCheck.AutoSize = $true
        [void]$basic.Controls.Add($kalmanCheck)
        $kalmanCheck.add_CheckedChanged({
            $kalmanQBox.Enabled = $kalmanCheck.Checked
            $kalmanRBox.Enabled = $kalmanCheck.Checked
            $adaptiveKalmanCheck.Enabled = $kalmanCheck.Checked
        })

        $adaptiveKalmanCheck = New-Object System.Windows.Forms.CheckBox
        $adaptiveKalmanCheck.Text = '自適應卡曼（IAE 調整 R；Q 依運動狀態調整）'
        $adaptiveKalmanCheck.Location = New-Object System.Drawing.Point(44, 186)
        $adaptiveKalmanCheck.AutoSize = $true
        $adaptiveKalmanCheck.Enabled = $false
        [void]$basic.Controls.Add($adaptiveKalmanCheck)

        $kalmanQLabel = New-Object System.Windows.Forms.Label
        $kalmanQLabel.Text = '卡曼 Q（程序雜訊強度）'
        $kalmanQLabel.Location = New-Object System.Drawing.Point(24, 218)
        $kalmanQLabel.AutoSize = $true
        [void]$basic.Controls.Add($kalmanQLabel)
        $kalmanQBox = New-Object System.Windows.Forms.TextBox
        $kalmanQBox.Text = '自動'
        $kalmanQBox.Location = New-Object System.Drawing.Point(190, 215)
        $kalmanQBox.Width = 150
        $kalmanQBox.Enabled = $false
        [void]$basic.Controls.Add($kalmanQBox)

        $kalmanRLabel = New-Object System.Windows.Forms.Label
        $kalmanRLabel.Text = '卡曼 R（位置量測變異數）'
        $kalmanRLabel.Location = New-Object System.Drawing.Point(365, 218)
        $kalmanRLabel.AutoSize = $true
        [void]$basic.Controls.Add($kalmanRLabel)
        $kalmanRBox = New-Object System.Windows.Forms.TextBox
        $kalmanRBox.Text = '自動'
        $kalmanRBox.Location = New-Object System.Drawing.Point(540, 215)
        $kalmanRBox.Width = 140
        $kalmanRBox.Enabled = $false
        [void]$basic.Controls.Add($kalmanRBox)

        $kalmanHint = New-Object System.Windows.Forms.Label
        $kalmanHint.Text = 'Q、R 欄位設定基準值；留空或「自動」依資料估算。自適應後的逐筆 Q、R 另列於報告。'
        $kalmanHint.Location = New-Object System.Drawing.Point(24, 242)
        $kalmanHint.AutoSize = $true
        [void]$basic.Controls.Add($kalmanHint)

        $runButton = New-Object System.Windows.Forms.Button
        $runButton.Text = '選取 Excel 並開始分析'
        $runButton.Location = New-Object System.Drawing.Point(24, 266)
        $runButton.AutoSize = $true
        [void]$basic.Controls.Add($runButton)
        $ncLabel = New-Object System.Windows.Forms.Label
        $ncLabel.Text = 'NC Code（可直接貼上，將逐行加入報告工作表）'
        $ncLabel.Location = New-Object System.Drawing.Point(24, 306)
        $ncLabel.AutoSize = $true
        [void]$basic.Controls.Add($ncLabel)
        $ncCodeBox = New-Object System.Windows.Forms.TextBox
        $ncCodeBox.Multiline = $true
        $ncCodeBox.AcceptsReturn = $true
        $ncCodeBox.AcceptsTab = $true
        $ncCodeBox.ScrollBars = 'Both'
        $ncCodeBox.WordWrap = $false
        $ncCodeBox.Location = New-Object System.Drawing.Point(24, 331)
        $ncCodeBox.Size = New-Object System.Drawing.Size(660, 149)
        [void]$basic.Controls.Add($ncCodeBox)
        $status = New-Object System.Windows.Forms.Label
        $status.Text = '選取 .xlsx 或 .xlsm 檔案開始分析。'
        $status.Location = New-Object System.Drawing.Point(24, 500)
        $status.Size = New-Object System.Drawing.Size(660, 70)
        [void]$basic.Controls.Add($status)

        $aboutText = New-Object System.Windows.Forms.Label
        $aboutText.Text = "版本：$env:AXIS_SCURVE_VERSION`r`n`r`nX 軸 α–β 濾波與 S 曲線分析`r`n`r`n選取包含 X 軸命令位置、回授位置、速度與 tick 資料的 Excel 活頁簿，產生包含分析總覽、S 曲線分段、門檻明細、圖表及方法設定的報告。`r`n`r`n卡曼濾波可將估測位置、速度與加速度加入逐筆估算工作表及圖表。自適應卡曼功能說明：`r`nR：以 IAE 依創新量更新，更新率 0.05，限制在基準 R 的 0.1–10 倍。`r`nQ：依命令速度判定運動狀態；運動時使用基準 Q，靜止時使用基準 Q 的 0.1 倍。兩者不以同一誤差訊號共同自由最佳化。`r`nExcel 逐筆估算頁新增每筆採用的 Q、R；方法與設定頁記錄調適方式。命令列支援 --adaptive-kalman，且必須搭配 --kalman-filter。此功能為離線分析，不直接控制 CNC。`r`n`r`n支援格式：.xlsx、.xlsm`r`n報告會儲存於來源 Excel 同一資料夾。`r`n`r`n開始者 Mail：yurenglee@gmail.com"
        $aboutText.Location = New-Object System.Drawing.Point(24, 24)
        $aboutText.Size = New-Object System.Drawing.Size(660, 580)
        [void]$about.Controls.Add($aboutText)
        Set-ControlFontSize $form 12

        $script:analysisProcess = $null
        $script:analysisTimer = New-Object System.Windows.Forms.Timer
        $script:analysisTimer.Interval = 300
        $script:analysisTimer.add_Tick({
            if (-not $script:analysisProcess -or -not $script:analysisProcess.HasExited) {
                return
            }
            $script:analysisTimer.Stop()
            $stdout = $script:analysisProcess.StandardOutput.ReadToEnd()
            $stderr = $script:analysisProcess.StandardError.ReadToEnd()
            $exitCode = $script:analysisProcess.ExitCode
            $script:analysisProcess.Dispose()
            $script:analysisProcess = $null
            if ($script:ncCodeTempPath -and [System.IO.File]::Exists($script:ncCodeTempPath)) {
                [System.IO.File]::Delete($script:ncCodeTempPath)
                $script:ncCodeTempPath = $null
            }
            $runButton.Enabled = $true
            $resultText = ($stdout + "`r`n" + $stderr).Trim()
            if ($exitCode -ne 0) {
                $status.Text = '分析未完成。'
                [void][System.Windows.Forms.MessageBox]::Show(
                    $resultText, '分析失敗',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            } else {
                $noiseMatch = [regex]::Match(
                    $stdout, 'Kalman noise parameters: Q=([^,\s]+), R=([^\s]+)')
                if ($noiseMatch.Success) {
                    $kalmanQBox.Text = $noiseMatch.Groups[1].Value
                    $kalmanRBox.Text = $noiseMatch.Groups[2].Value
                }
                $status.Text = '分析完成。' + "`r`n" + $resultText
                [void][System.Windows.Forms.MessageBox]::Show(
                    $resultText, '分析完成',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        })

        $runButton.add_Click({
            $culture = [System.Globalization.CultureInfo]::InvariantCulture
            $styles = [System.Globalization.NumberStyles]::Float
            $alpha = 0.0
            $beta = 0.0
            $validAlpha = [double]::TryParse($alphaBox.Text, $styles, $culture, [ref]$alpha)
            $validBeta = [double]::TryParse($betaBox.Text, $styles, $culture, [ref]$beta)
            if (-not $autoCheck.Checked -and
                (-not $validAlpha -or -not $validBeta -or
                 [double]::IsNaN($alpha) -or [double]::IsInfinity($alpha) -or
                 [double]::IsNaN($beta) -or [double]::IsInfinity($beta) -or
                 $alpha -le 0 -or $alpha -ge 2 -or $beta -le 0 -or $beta -ge (4 - 2 * $alpha))) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    '請輸入有效數值，並符合 0 < α < 2 及 0 < β < 4−2α。',
                    '參數錯誤', [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $kalmanQ = 0.0
            $kalmanR = 0.0
            $qText = $kalmanQBox.Text.Trim()
            $rText = $kalmanRBox.Text.Trim()
            $qAutomatic = [string]::IsNullOrWhiteSpace($qText) -or $qText -eq '自動'
            $rAutomatic = [string]::IsNullOrWhiteSpace($rText) -or $rText -eq '自動'
            if ($kalmanCheck.Checked -and
                ((-not $qAutomatic -and
                  (-not [double]::TryParse($qText, $styles, $culture, [ref]$kalmanQ) -or
                   [double]::IsNaN($kalmanQ) -or [double]::IsInfinity($kalmanQ) -or $kalmanQ -lt 0)) -or
                 (-not $rAutomatic -and
                  (-not [double]::TryParse($rText, $styles, $culture, [ref]$kalmanR) -or
                   [double]::IsNaN($kalmanR) -or [double]::IsInfinity($kalmanR) -or $kalmanR -le 0)))) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    'Q 必須是有限的非負數，R 必須是有限的正數；留空或輸入「自動」可使用資料估算值。',
                    '卡曼參數錯誤', [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = '選取 X 軸量測 Excel'
            $dialog.Filter = 'Excel 活頁簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm'
            $dialog.CheckFileExists = $true
            if ($dialog.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) {
                $dialog.Dispose()
                return
            }
            $file = $dialog.FileName
            $dialog.Dispose()
            $output = [System.IO.Path]::Combine(
                [System.IO.Path]::GetDirectoryName($file),
                [System.IO.Path]::GetFileNameWithoutExtension($file) + '_scurve_analysis.xlsx')
            if ([System.IO.File]::Exists($output) -and
                [System.Windows.Forms.MessageBox]::Show(
                    "分析報告已存在：`r`n$output`r`n`r`n要覆寫嗎？",
                    '確認覆寫', [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question) -ne
                    [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
            $runButton.Enabled = $false
            $status.Text = '正在分析：' + [System.IO.Path]::GetFileName($file) + '…'
            $form.Refresh()
            try {
                $info = New-Object System.Diagnostics.ProcessStartInfo
                $info.FileName = $env:AXIS_SCURVE_PYTHON
                $info.Arguments = (Quote-Argument $env:AXIS_SCURVE_SCRIPT) + ' ' +
                    (Quote-Argument $file) + ' --overwrite'
                if ($autoCheck.Checked) {
                    $info.Arguments += ' --auto-alpha-beta'
                } else {
                    $info.Arguments += ' --alpha ' + $alpha.ToString('R', $culture) +
                        ' --beta ' + $beta.ToString('R', $culture)
                }
                if ($kalmanCheck.Checked) {
                    $info.Arguments += ' --kalman-filter'
                    if ($adaptiveKalmanCheck.Checked) {
                        $info.Arguments += ' --adaptive-kalman'
                    }
                    if (-not $qAutomatic) {
                        $info.Arguments += ' --kalman-q ' + $kalmanQ.ToString('R', $culture)
                    }
                    if (-not $rAutomatic) {
                        $info.Arguments += ' --kalman-r ' + $kalmanR.ToString('R', $culture)
                    }
                }
                if ($ncCodeBox.Text.Length -gt 0) {
                    $script:ncCodeTempPath = [System.IO.Path]::GetTempFileName()
                    [System.IO.File]::WriteAllText(
                        $script:ncCodeTempPath, $ncCodeBox.Text,
                        [System.Text.UTF8Encoding]::new($false))
                    $info.Arguments += ' --nc-code-file ' + (Quote-Argument $script:ncCodeTempPath)
                }
                $info.UseShellExecute = $false
                $info.CreateNoWindow = $true
                $info.RedirectStandardOutput = $true
                $info.RedirectStandardError = $true
                $script:analysisProcess = New-Object System.Diagnostics.Process
                $script:analysisProcess.StartInfo = $info
                if (-not $script:analysisProcess.Start()) {
                    throw '無法啟動 Python 分析程序。'
                }
                $script:analysisTimer.Start()
            } catch {
                if ($script:analysisProcess) {
                    $script:analysisProcess.Dispose()
                    $script:analysisProcess = $null
                }
                if ($script:ncCodeTempPath -and [System.IO.File]::Exists($script:ncCodeTempPath)) {
                    [System.IO.File]::Delete($script:ncCodeTempPath)
                    $script:ncCodeTempPath = $null
                }
                $script:analysisTimer.Stop()
                $runButton.Enabled = $true
                $status.Text = '分析未完成。'
                [void][System.Windows.Forms.MessageBox]::Show(
                    $_.Exception.Message, '分析失敗',
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })

        [void]$form.ShowDialog()
        $script:analysisTimer.Stop()
        $script:analysisTimer.Dispose()
        if ($script:analysisProcess) {
            $script:analysisProcess.Dispose()
            $script:analysisProcess = $null
        }
        if ($script:ncCodeTempPath -and [System.IO.File]::Exists($script:ncCodeTempPath)) {
            [System.IO.File]::Delete($script:ncCodeTempPath)
            $script:ncCodeTempPath = $null
        }
        $form.Dispose()
    """
    powershell = (Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32"
                  / "WindowsPowerShell" / "v1.0" / "powershell.exe")
    if not powershell.is_file():
        raise RuntimeError("找不到 Windows PowerShell，無法啟動圖形介面。")
    environment = os.environ.copy()
    python_executable = Path(sys.executable)
    if python_executable.stem.casefold() == "pythonw":
        console_python = python_executable.with_name("python.exe")
        if console_python.is_file():
            python_executable = console_python
    environment["AXIS_SCURVE_PYTHON"] = str(python_executable)
    environment["AXIS_SCURVE_SCRIPT"] = str(Path(__file__).resolve())
    environment["AXIS_SCURVE_VERSION"] = APP_VERSION
    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8-sig", suffix=".ps1", prefix="axis_scurve_gui_", delete=False
    ) as temporary:
        gui_script = Path(temporary.name)
        temporary.write(script)
    environment["AXIS_SCURVE_GUI_SCRIPT"] = str(gui_script)
    launcher = "Invoke-Expression ([System.IO.File]::ReadAllText($env:AXIS_SCURVE_GUI_SCRIPT))"
    encoded = base64.b64encode(launcher.encode("utf-16-le")).decode("ascii")
    try:
        finished = subprocess.run(
            [str(powershell), "-NoLogo", "-NoProfile", "-STA", "-EncodedCommand", encoded],
            stdin=subprocess.DEVNULL, env=environment, check=False,
        )
    finally:
        gui_script.unlink(missing_ok=True)
    if finished.returncode:
        raise RuntimeError("Windows 圖形介面啟動失敗，請確認 Windows Forms 可正常使用。")


def launch_gui():
    if sys.platform == "win32":
        launch_windows_gui()
        return

    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("X 軸 α–β 濾波與 S 曲線分析")
    root.geometry("720x770")
    root.minsize(650, 640)
    root.option_add("*Font", "TkDefaultFont")
    for font_name in tkfont.names(root):
        tkfont.nametofont(font_name, root=root).configure(size=DEFAULT_FONT_SIZE)

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True, padx=12, pady=12)

    basic = ttk.Frame(notebook, padding=20)
    notebook.add(basic, text="基本功能區")
    basic.columnconfigure(1, weight=1)
    basic.columnconfigure(3, weight=1)

    ttk.Label(basic, text="字體大小（空白使用 12）").grid(row=0, column=0, sticky="w")
    font_size_entry = ttk.Entry(basic, width=8)
    font_size_entry.grid(row=0, column=1, sticky="w", padx=(12, 0))

    def apply_font_size():
        value = font_size_entry.get().strip()
        if not value:
            size = DEFAULT_FONT_SIZE
        else:
            try:
                size = int(value)
            except ValueError:
                messagebox.showerror("字體大小無效", "請輸入 6 到 72 的整數；留空則使用 12。", parent=root)
                return
            if not 6 <= size <= 72:
                messagebox.showerror("字體大小無效", "請輸入 6 到 72 的整數；留空則使用 12。", parent=root)
                return
        for font_name in tkfont.names(root):
            tkfont.nametofont(font_name, root=root).configure(size=size)

    ttk.Button(basic, text="變更", command=apply_font_size).grid(
        row=0, column=2, sticky="w", padx=(8, 0)
    )

    ttk.Label(basic, text="α（位置更新增益）").grid(row=1, column=0, sticky="w", pady=8)
    alpha_entry = ttk.Entry(basic, width=18)
    alpha_entry.insert(0, "0.64")
    alpha_entry.grid(row=1, column=1, sticky="w", padx=(12, 0), pady=8)

    ttk.Label(basic, text="β（速度更新增益）").grid(row=2, column=0, sticky="w", pady=8)
    beta_entry = ttk.Entry(basic, width=18)
    beta_entry.insert(0, "0.16")
    beta_entry.grid(row=2, column=1, sticky="w", padx=(12, 0), pady=8)

    auto_alpha_beta = tk.BooleanVar(value=False)

    def toggle_manual_parameters():
        state = "disabled" if auto_alpha_beta.get() else "normal"
        alpha_entry.configure(state=state)
        beta_entry.configure(state=state)

    ttk.Checkbutton(
        basic,
        text="自調適 α–β（以 Vxfb RMSE 最小化）",
        variable=auto_alpha_beta,
        command=toggle_manual_parameters,
    ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

    run_kalman = tk.BooleanVar(value=False)
    adaptive_kalman = tk.BooleanVar(value=False)

    def toggle_kalman_parameters():
        state = "normal" if run_kalman.get() else "disabled"
        kalman_q_entry.configure(state=state)
        kalman_r_entry.configure(state=state)
        adaptive_kalman_check.configure(state=state)

    ttk.Checkbutton(
        basic,
        text="啟用卡曼濾波並輸出估測位置、速度與加速度",
        variable=run_kalman,
        command=toggle_kalman_parameters,
    ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

    adaptive_kalman_check = ttk.Checkbutton(
        basic,
        text="自適應卡曼（IAE 調整 R；Q 依運動狀態調整）",
        variable=adaptive_kalman,
        state="disabled",
    )
    adaptive_kalman_check.grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))

    ttk.Label(basic, text="卡曼 Q（程序雜訊強度）").grid(
        row=6, column=0, sticky="w", pady=(6, 0)
    )
    kalman_q_entry = ttk.Entry(basic, width=18)
    kalman_q_entry.insert(0, "自動")
    kalman_q_entry.configure(state="disabled")
    kalman_q_entry.grid(row=6, column=1, sticky="w", padx=(12, 0), pady=(6, 0))

    ttk.Label(basic, text="卡曼 R（位置量測變異數）").grid(
        row=6, column=2, sticky="w", padx=(16, 0), pady=(6, 0)
    )
    kalman_r_entry = ttk.Entry(basic, width=18)
    kalman_r_entry.insert(0, "自動")
    kalman_r_entry.configure(state="disabled")
    kalman_r_entry.grid(row=6, column=3, sticky="w", padx=(8, 0), pady=(6, 0))

    ttk.Label(
        basic,
        text="Q、R 欄位設定基準值；留空或「自動」依資料估算。自適應後的逐筆 Q、R 另列於報告。",
        wraplength=660,
    ).grid(row=7, column=0, columnspan=5, sticky="w", pady=(6, 0))

    ttk.Label(basic, text="NC Code（貼上後逐行加入報告工作表）").grid(
        row=8, column=0, columnspan=4, sticky="w", pady=(12, 4)
    )
    nc_code = tk.Text(basic, height=12, wrap="none")
    nc_code.grid(row=9, column=0, columnspan=4, sticky="nsew")
    nc_scroll = ttk.Scrollbar(basic, orient="vertical", command=nc_code.yview)
    nc_scroll.grid(row=9, column=4, sticky="ns")
    nc_code.configure(yscrollcommand=nc_scroll.set)
    basic.rowconfigure(9, weight=1)

    status = tk.StringVar(value="選取 .xlsx 或 .xlsm 檔案開始分析。")
    ttk.Label(basic, textvariable=status, wraplength=620).grid(
        row=12, column=0, columnspan=5, sticky="w", pady=(12, 8)
    )

    def analyze_file():
        if auto_alpha_beta.get():
            alpha, beta = 0.64, 0.16
        else:
            try:
                alpha = float(alpha_entry.get())
                beta = float(beta_entry.get())
            except ValueError:
                messagebox.showerror("參數錯誤", "α 與 β 必須是有效數字。", parent=root)
                return
            if not all(map(math.isfinite, (alpha, beta))) or not 0 < alpha < 2 or not 0 < beta < 4 - 2 * alpha:
                messagebox.showerror("參數錯誤", "請符合 0 < α < 2 及 0 < β < 4−2α。", parent=root)
                return
        kalman_q = None
        kalman_r = None
        if run_kalman.get():
            try:
                q_text = kalman_q_entry.get().strip()
                r_text = kalman_r_entry.get().strip()
                kalman_q = None if not q_text or q_text == "自動" else float(q_text)
                kalman_r = None if not r_text or r_text == "自動" else float(r_text)
            except ValueError:
                messagebox.showerror("卡曼參數錯誤", "Q、R 必須是有效數字或「自動」。", parent=root)
                return
            if ((kalman_q is not None and (not math.isfinite(kalman_q) or kalman_q < 0))
                    or (kalman_r is not None and (not math.isfinite(kalman_r) or kalman_r <= 0))):
                messagebox.showerror(
                    "卡曼參數錯誤",
                    "Q 必須是有限的非負數，R 必須是有限的正數。",
                    parent=root,
                )
                return

        filename = filedialog.askopenfilename(
            parent=root,
            title="選取 X 軸量測 Excel",
            filetypes=[("Excel 活頁簿", "*.xlsx *.xlsm"), ("所有檔案", "*.*")],
        )
        if not filename:
            return

        source = Path(filename).expanduser().resolve()
        output = source.with_name(source.stem + "_scurve_analysis.xlsx")
        if output.exists() and not messagebox.askyesno(
            "確認覆寫", f"分析報告已存在：\n{output}\n\n要覆寫嗎？", parent=root
        ):
            return

        try:
            settings = Settings(1.0, alpha, beta, 3, "pulse/sample", 5, 8, 2, 5, 12)
            status.set(f"正在分析：{source.name}…")
            root.update_idletasks()
            log = load_log(source, None, settings)
            adaptive_rmse = None
            if auto_alpha_beta.get():
                settings, adaptive_rmse = tune_alpha_beta(log, settings)
            estimate = alpha_beta(log, settings)
            kalman_estimate = (
                kalman_filter(
                    log, kalman_q, kalman_r, adaptive_kalman.get(),
                    settings.motion_threshold,
                )
                if run_kalman.get() else None
            )
            moves = detect_motions(log, settings)
            ramps = analyze_ramps(log, estimate, moves, settings)
            write_report(output, source, log, estimate, moves, ramps, settings,
                         nc_code.get("1.0", "end-1c"), adaptive_rmse, kalman_estimate)
        except (OSError, ValueError, BadZipFile, openpyxl.utils.exceptions.InvalidFileException,
                xlsxwriter.exceptions.XlsxWriterException) as exc:
            status.set("分析未完成。")
            messagebox.showerror("分析失敗", str(exc), parent=root)
            return

        adaptive_summary = (
            f"\n自調適結果：α={settings.alpha:.6g}、β={settings.beta:.6g}、"
            f"Vxfb RMSE={adaptive_rmse:.6g} pulse/ms。"
            if adaptive_rmse is not None else ""
        )
        kalman_summary = ""
        if kalman_estimate is not None:
            kalman_q_entry.configure(state="normal")
            kalman_r_entry.configure(state="normal")
            kalman_q_entry.delete(0, "end")
            kalman_q_entry.insert(0, f"{kalman_estimate.process_noise:.8g}")
            kalman_r_entry.delete(0, "end")
            kalman_r_entry.insert(0, f"{kalman_estimate.measurement_variance:.8g}")
            if not run_kalman.get():
                toggle_kalman_parameters()
            kalman_summary = (
                f"\n卡曼 Q={kalman_estimate.process_noise:.8g}，"
                f"R={kalman_estimate.measurement_variance:.8g}。"
            )
            if kalman_estimate.adaptive:
                kalman_summary += (
                    f"\n自適應範圍：Q={np.min(kalman_estimate.process_noise_values):.8g}～"
                    f"{np.max(kalman_estimate.process_noise_values):.8g}，"
                    f"R={np.min(kalman_estimate.measurement_variance_values):.8g}～"
                    f"{np.max(kalman_estimate.measurement_variance_values):.8g}。"
                )
        status.set(f"分析完成：共 {len(log.time):,} 筆資料、{len(moves)} 段運動。"
                   f"{adaptive_summary}{kalman_summary}\n報告：{output}")
        messagebox.showinfo(
            "分析完成", f"已產生分析報告：\n{output}{adaptive_summary}{kalman_summary}", parent=root
        )

    ttk.Button(basic, text="選取 Excel 並開始分析", command=analyze_file).grid(
        row=11, column=0, columnspan=2, sticky="w", pady=(10, 0)
    )

    about = ttk.Frame(notebook, padding=20)
    notebook.add(about, text="軟體功能說明")
    ttk.Label(
        about,
        text=(
            f"版本：{APP_VERSION}\n\n"
            "X 軸 α–β 濾波與 S 曲線分析\n\n"
            "選取包含 X 軸命令位置、回授位置、速度與 tick 資料的 Excel 活頁簿，"
            "產生包含分析總覽、S 曲線分段、門檻明細、圖表及方法設定的報告。"
            "可選擇執行卡曼濾波，將估測位置、速度與加速度加入逐筆估算工作表及圖表；"
            "Q、R 為基準雜訊參數。自適應卡曼以 IAE 調整 R，並依命令速度運動狀態調整 Q。"
            "此功能為離線分析，不直接控制 CNC。\n\n"
            "自適應卡曼參數與輸出：\n"
            "R：以 IAE 依創新量更新，更新率 0.05，限制在基準 R 的 0.1–10 倍。\n"
            "Q：依命令速度判定運動狀態；運動時使用基準 Q，靜止時使用基準 Q 的 0.1 倍。"
            "兩者不以同一誤差訊號共同自由最佳化。\n"
            "Excel 逐筆估算頁新增每筆採用的 Q、R；方法與設定頁記錄調適方式。"
            "命令列支援 --adaptive-kalman，且必須搭配 --kalman-filter。\n\n"
            "支援格式：.xlsx、.xlsm\n"
            "報告會儲存於來源 Excel 同一資料夾。\n\n"
            "開始者 Mail：yurenglee@gmail.com"
        ),
        justify="left",
        wraplength=540,
    ).pack(anchor="nw")

    root.mainloop()


def main():
    if len(sys.argv) == 1:
        launch_gui()
        return 0

    parser = args_parser()
    options = parser.parse_args()
    if not all(map(math.isfinite, (options.tick_ms, options.alpha, options.beta, options.motion_threshold))):
        parser.error("時間、增益與運動門檻須為有限數值。")
    if options.tick_ms <= 0 or not 0 < options.alpha < 2 or not 0 < options.beta < 4 - 2 * options.alpha:
        parser.error("須符合 tick-ms > 0、0 < α < 2 及 0 < β < 4−2α。")
    if (min(options.half_window, options.min_motion, options.crossing_samples) < 1
            or min(options.motion_threshold, options.merge_gap, options.chart_ramps) < 0):
        parser.error("半窗口、最短運動和門檻確認須至少為 1。速度門檻、合併間隔和圖數須為非負值。")
    if options.kalman_q is not None and (not math.isfinite(options.kalman_q) or options.kalman_q < 0):
        parser.error("卡曼 Q 須為有限的非負數值。")
    if options.kalman_r is not None and (not math.isfinite(options.kalman_r) or options.kalman_r <= 0):
        parser.error("卡曼 R 須為有限的正數。")
    if not options.kalman_filter and (
            options.kalman_q is not None or options.kalman_r is not None or options.adaptive_kalman):
        parser.error("--kalman-q、--kalman-r 與 --adaptive-kalman 需搭配 --kalman-filter 使用。")

    try:
        chosen = options.input or select_input(options.no_dialog)
        if chosen is None:
            print("已取消選檔。")
            return 0
        source = chosen.expanduser().resolve()
        if not source.is_file() or source.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("請指定存在的 .xlsx 或 .xlsm 來源檔。舊版 .xls 可在 Excel 另存為 .xlsx。")
        output = (options.output or source.with_name(source.stem + "_scurve_analysis.xlsx")).expanduser().resolve()
        if output.suffix.lower() != ".xlsx" or output == source or (output.exists() and output.samefile(source)):
            raise ValueError("輸出必須是與來源不同的 .xlsx 檔案。")
        if output.exists() and not options.overwrite:
            raise ValueError(f"報告已存在：{output}\n請使用 -o 指定其他路徑，或加 --overwrite 更新該報告。")
        if not output.parent.is_dir():
            raise ValueError(f"找不到輸出資料夾：{output.parent}")

        settings = Settings(options.tick_ms, options.alpha, options.beta, options.half_window,
                            options.velocity_unit, options.motion_threshold, options.min_motion,
                            options.merge_gap, options.crossing_samples, options.chart_ramps)
        nc_code = options.nc_code_file.read_text(encoding="utf-8-sig") if options.nc_code_file else ""
        print(f"讀取 {source.name}…", flush=True)
        log = load_log(source, options.sheet, settings)
        adaptive_rmse = None
        if options.auto_alpha_beta:
            settings, adaptive_rmse = tune_alpha_beta(log, settings)
            print(f"自動調適結果：α={settings.alpha:.6g}，β={settings.beta:.6g}，"
                  f"Vxfb RMSE={adaptive_rmse:.6g} pulse/ms。", flush=True)
        estimate = alpha_beta(log, settings)
        kalman_estimate = (
            kalman_filter(
                log, options.kalman_q, options.kalman_r, options.adaptive_kalman,
                settings.motion_threshold,
            )
            if options.kalman_filter else None
        )
        if kalman_estimate is not None:
            print(
                f"Kalman noise parameters: Q={kalman_estimate.process_noise:.17g}, "
                f"R={kalman_estimate.measurement_variance:.17g}",
                flush=True,
            )
            if kalman_estimate.adaptive:
                q_values = kalman_estimate.process_noise_values
                r_values = kalman_estimate.measurement_variance_values
                print(
                    f"Adaptive Kalman ranges: Q={np.min(q_values):.8g}..{np.max(q_values):.8g}, "
                    f"R={np.min(r_values):.8g}..{np.max(r_values):.8g}",
                    flush=True,
                )
        moves = detect_motions(log, settings)
        ramps = analyze_ramps(log, estimate, moves, settings)
        print(f"共 {len(log.time):,} 筆、{len(moves)} 段運動。正在產生報告…", flush=True)
        write_report(output, source, log, estimate, moves, ramps, settings, nc_code,
                     adaptive_rmse, kalman_estimate)
        print(f"已產生：{output}")
        return 0
    except (OSError, ValueError, BadZipFile, openpyxl.utils.exceptions.InvalidFileException,
            xlsxwriter.exceptions.XlsxWriterException) as exc:
        print(f"無法完成分析：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
