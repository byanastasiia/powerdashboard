# noinspection package-requirements
"""
Дашборд анализа энергопотребления скважин
==========================================
Установка:  pip install dash plotly pandas openpyxl
Запуск:     python dashboard.py

Ожидаемая структура Excel-файла (лист "data"):
───────────────────────────────────────────────────────────
| timestamp           | field        | well  | electricity_fact | electricity_plan | liquid | oil  | ure_fact | ure_plan |
| 2024-05-01 08:00:00 | Месторождение А | 2101 | 55000           | 53000           | 3200   | 420  | 17.2    | 16.5     |
───────────────────────────────────────────────────────────
Колонки:
  timestamp        — datetime, метка времени замера
  field            — название месторождения (для фильтра)
  well             — номер скважины
  electricity_fact — фактическое потребление, кВт·ч
  electricity_plan — расчётное потребление, кВт·ч
  liquid           — добыча жидкости, м³
  oil              — добыча нефти, т
  ure_fact         — УРЭ факт, кВт·ч/м³
  ure_plan         — УРЭ расч., кВт·ч/м³
  mode             — режим скважины: "в работе" | "в накоплении" | "в простое" | "в бездействии"
"""

import os
import json
import base64
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

import dash
from dash import dcc, html, Input, Output, State, callback_context, ALL, dash_table
import dash_bootstrap_components as dbc
from dash.exceptions import PreventUpdate
import plotly.graph_objects as go

import io

from dash import Dash

# ════════════════════════════════════════════════
#  ПУТЬ К ФАЙЛУ — поменяйте на свой
# ════════════════════════════════════════════════
EXCEL_PATH = "data.xlsx"        # абсолютный или относительный путь
EXCEL_SHEET = "data"            # имя листа

# ── Мероприятия по снижению УРЭ ───────────────────
# Excel с мероприятиями кладите (или загружайте через интерфейс, вкладка
# «Рейтинг скважин») сюда — файл автоматически подхватится дашбордом
# и будет доступен по прямой ссылке /meropriyatiya.xlsx
MEROPRIYATIYA_DIR  = "assets"
MEROPRIYATIYA_PATH = os.path.join(MEROPRIYATIYA_DIR, "meropriyatiya.xlsx")
MEROPRIYATIYA_COLUMNS = [
    "Скважина", "Мероприятие", "Рекомендации", "Возможная экономия электроэнергии, кВт",
]
ADDITIONAL_DATA_DIR  = "data_uploads"          # НЕ assets — не должно быть публично доступно
ADDITIONAL_DATA_PATH = os.path.join(ADDITIONAL_DATA_DIR, "additional_consumption.xlsx")

# ── Цвета (палитра — по референсу) ────────────────
GREEN_DARK   = "rgba(34, 197, 94, 1)"     # #22C55E — «норма»
GREEN_MID    = "rgba(34, 197, 94, 0.7)"
GREEN_LIGHT  = "rgba(34, 197, 94, 0.5)"
GREEN_TRANSP = "rgba(34, 197, 94, 0.3)"
HERBAL       = "#9BBD1E"
GREEN        = "rgba(16, 185, 129, 1)"    # #10B981 — доп. оттенок зелёного (для тепловой карты)
YELLOW       = "rgba(245, 158, 11, 1)"    # #F59E0B — «внимание» (amber)
YELLOW_LIGHT = "rgba(245, 158, 11, 0.5)"
YELLOW_TRANSP = "rgba(245, 158, 11, 0.3)"
ORANGE        = "rgba(239, 68, 68, 1)"    # #EF4444 — «критично» / акцент отклонения (красный)
ORANGE_MID    = "rgba(239, 68, 68, 0.7)"
ORANGE_LIGHT  = "rgba(239, 68, 68, 0.5)"
ORANGE_TRANSP = "rgba(239, 68, 68, 0.3)"
BLUE         = "rgba(79, 124, 255, 1)"    # #4F7CFF — основной / инфо
BLUE_LIGHT   = "rgba(79, 124, 255, 0.5)"
PURPLE       = "rgba(139, 92, 246, 1)"    # #8B5CF6 — энергопотребление
TEAL         = "rgba(6, 182, 212, 1)"     # #06B6D4 — жидкость/доп. акцент
GREY_DARK    = "#7B9AA8"
GREY_MID     = '#B6C7CF'
GREY_LIGHT   = "#b0bec5"
GREY_TRANSP = "#E2E9EC"
BG           = "#f5f6fa"      # светлый (почти белый) фон страницы
CARD_BG      = "#ffffff"      # белые карточки
TEXT_DARK    = "#14171c"      # чёрный текст (основной)
TEXT_MUTED   = "#6b7280"      # приглушённый серый для подписей
BORDER_SOFT  = "#e9eaef"

# ── общий стиль вкладок (dcc.Tabs) для всего дашборда ──
TAB_STYLE = {"padding": "8px 4px", "border": "none", "borderBottom": "2px solid transparent",
             "fontSize": "13px", "color": TEXT_MUTED, "fontWeight": "600"}
TAB_SELECTED_STYLE = {**TAB_STYLE, "borderBottom": f"2px solid {GREEN_DARK}", "color": TEXT_DARK}


def with_alpha(color_str: str, a: float) -> str:
    """Возвращает тот же цвет с новой прозрачностью — для полупрозрачных блоков.
    Понимает и rgba(...)/rgb(...), и HEX (#rrggbb)."""
    color_str = color_str.strip()
    if color_str.startswith("#"):
        hex_str = color_str.lstrip("#")
        if len(hex_str) == 3:
            hex_str = "".join(c * 2 for c in hex_str)
        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)
        return f"rgba({r}, {g}, {b}, {a})"
    inner = color_str[color_str.find("(") + 1: color_str.find(")")]
    r, g, b = [p.strip() for p in inner.split(",")[:3]]
    return f"rgba({r}, {g}, {b}, {a})"


def glow(rgba_str: str, a: float = 0.18) -> str:
    """Мягкая цветная тень под карточкой в стиле референса."""
    return f"0 6px 20px {with_alpha(rgba_str, a)}, 0 1px 3px rgba(20,23,28,0.04)"


CARD_STYLE = {
    "background": CARD_BG, "border": f"1px solid {BORDER_SOFT}", "borderRadius": "16px",
    "padding": "16px", "boxShadow": glow(GREY_DARK, 0.08),
}

MODES = ["в работе", "в накоплении", "в простое", "в бездействии"]
MODE_COLORS = {
    "в работе":      GREEN_DARK,
    "в накоплении":  HERBAL,
    "в простое":     YELLOW,
    "в бездействии": ORANGE,
}

PERIODS = {
    "8ч":  timedelta(hours=8),
    "24ч": timedelta(hours=24),
    "7д":  timedelta(days=7),
    "1м":  timedelta(days=30),
}

# ════════════════════════════════════════════════
#  МИНИМАЛИСТИЧНЫЕ SVG-ИКОНКИ
# ════════════════════════════════════════════════

import urllib.parse as _urlparse

ICON_PATHS = {
    "bolt":        '<path d="M13 2 4 14h6l-1 8 9-12h-6l1-8z"/>',
    "plug":        '<circle cx="12" cy="12" r="8"/><circle cx="12" cy="12" r="2.4" fill="currentColor" stroke="none"/>',
    "droplet":     '<path d="M12 3c4 5 7 8.6 7 12.2A7 7 0 1 1 5 15.2C5 11.6 8 8 12 3z"/>',
    "barrel":      '<rect x="6" y="4" width="12" height="16" rx="3"/><line x1="6" y1="10" x2="18" y2="10"/><line x1="6" y1="14" x2="18" y2="14"/>',
    "gauge":       '<path d="M4 16a8 8 0 0 1 16 0"/><line x1="12" y1="16" x2="16" y2="10"/><circle cx="12" cy="16" r="1.3" fill="currentColor" stroke="none"/>',
    "clipboard":   '<rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2"/><line x1="8" y1="10.5" x2="16" y2="10.5"/><line x1="8" y1="14.5" x2="16" y2="14.5"/>',
    "trend":       '<polyline points="4,17 9,10.5 13,13.5 20,5"/><polyline points="14,5 20,5 20,11"/>',
    "grid":        '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    "home":        '<path d="M4 11 12 4 20 11"/><path d="M6 10v9h5v-5h2v5h5v-9"/>',
    "search":      '<circle cx="10" cy="10" r="6.2"/><line x1="15" y1="15" x2="20.5" y2="20.5"/>',
    "close":       '<line x1="6" y1="6" x2="18" y2="18"/><line x1="18" y1="6" x2="6" y2="18"/>',
    "pin":         '<path d="M12 21s7-7.6 7-12.2A7 7 0 1 0 5 8.8C5 13.4 12 21 12 21z"/><circle cx="12" cy="8.8" r="2.4"/>',
    "chart-line":  '<polyline points="4,18 9,11 13,14 20,6"/>',
    "list":        '<line x1="5" y1="6" x2="19" y2="6"/><line x1="5" y1="12" x2="19" y2="12"/><line x1="5" y1="18" x2="19" y2="18"/>',
    "refresh":     '<path d="M4 12a8 8 0 0 1 13.7-5.7L20 8"/><path d="M20 4v4h-4"/><path d="M20 12a8 8 0 0 1-13.7 5.7L4 16"/><path d="M4 20v-4h4"/>',
    "activity":    '<polyline points="3,12 8,12 10,6 14,18 16,12 21,12"/>',
    "bulb":        '<path d="M9 18h6"/><path d="M10 21h4"/><path d="M12 3a6 6 0 0 0-3.4 10.9c.6.5 1 1.1 1.1 1.9v.2h4.6v-.2c0-.8.5-1.4 1.1-1.9A6 6 0 0 0 12 3z"/>',
    "upload":      '<path d="M12 3v12"/><polyline points="7,8 12,3 17,8"/><path d="M5 21h14"/>',
    "link":        '<path d="M9 15l6-6"/><path d="M7 10 4.6 12.4a4 4 0 0 0 5.7 5.7L13 15.7"/><path d="M17 14l2.4-2.4a4 4 0 0 0-5.7-5.7L11 8.3"/>',
    "chevron":     '<polyline points="9,6 15,12 9,18"/>',
}


def svg_icon(name: str, color: str = TEXT_DARK, size: int = 18, stroke_width: float = 1.8):
    """Строит минималистичную линейную SVG-иконку без внешних зависимостей."""
    path = ICON_PATHS.get(name, ICON_PATHS["bolt"])
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="{size}" height="{size}" '
        f'fill="none" stroke="{color}" stroke-width="{stroke_width}" '
        f'stroke-linecap="round" stroke-linejoin="round">{path}</svg>'
    )
    encoded = _urlparse.quote(svg)
    return html.Img(src=f"data:image/svg+xml;utf8,{encoded}",
                     style={"width": f"{size}px", "height": f"{size}px", "display": "block"})

# ════════════════════════════════════════════════
#  ЗАГРУЗКА И ГЕНЕРАЦИЯ ДЕМО-ДАННЫХ
# ════════════════════════════════════════════════

def generate_demo_data() -> pd.DataFrame:
    """Создаёт синтетический датасет, если Excel-файл не найден."""
    np.random.seed(42)
    now = datetime.now()
    fields = ["Месторождение А", "Месторождение Б", "Месторождение В"]
    wells_by_field = {
        "Месторождение А": ["1","2","6","2101","2103","2104","2108"],
        "Месторождение Б": ["2206","2210","2211","2402","2404","2405"],
        "Месторождение В": ["2406","2503","2504","2602","2702","2703","10р"],
    }

    # ── профиль каждой скважины: у всех разное характерное отклонение
    #    УРЭ факт от расчёта, поэтому цвета на тепловой карте различаются
    well_profiles = {}
    dev_pool = [1, 2, 4, 6, 8, 11, 14, 17, 21, 26, 31, 38]
    for field, wlist in wells_by_field.items():
        for w in wlist:
            well_profiles[w] = {
                "base_ure": np.random.uniform(7, 26),
                "dev":      np.random.choice(dev_pool),
            }

    rows = []
    # записи каждые 8 часов за 30 дней
    for days_back in range(30, -1, -1):
        for hour in [0, 8, 16]:
            ts = now - timedelta(days=days_back, hours=now.hour) + timedelta(hours=hour)
            for field, wlist in wells_by_field.items():
                for w in wlist:
                    base = np.random.uniform(40, 95)
                    prof = well_profiles[w]
                    ure_f = max(prof["base_ure"] + np.random.normal(0, 1.1), 1.0)
                    dev_now = prof["dev"] + np.random.normal(0, 1.8)
                    ure_p = max(ure_f / (1 + dev_now / 100), 0.5)
                    mode = np.random.choice(MODES, p=[0.55, 0.02, 0.03, 0.40])
                    rows.append({
                        "timestamp":        ts,
                        "field":            field,
                        "well":             w,
                        "electricity_fact": round(base * 1000),
                        "electricity_plan": round(base * 0.97 * 1000),
                        "liquid":           round(np.random.uniform(2000, 5000)),
                        "oil":              round(np.random.uniform(200, 600)),
                        "ure_fact":         round(ure_f, 2),
                        "ure_plan":         round(ure_p, 2),
                        "mode":             mode,
                    })
    return pd.DataFrame(rows)


import unicodedata

def _normalize_text_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Убирает невидимые пробелы/разницу кодировок в текстовых колонках."""
    for col in ("field", "well", "mode"):
        if col in df.columns:
            df[col] = (df[col].astype(str)
                       .str.strip()
                       .apply(lambda s: unicodedata.normalize("NFC", s)))
    return df


def load_data() -> pd.DataFrame:
    df = _load_base_data()

    if os.path.exists(ADDITIONAL_DATA_PATH):
        try:
            extra = pd.read_excel(ADDITIONAL_DATA_PATH, sheet_name=EXCEL_SHEET)
            extra.columns = extra.columns.str.strip().str.lower()
            extra["timestamp"] = pd.to_datetime(extra["timestamp"])
            extra["well"] = extra["well"].astype(str)   # ← нормализация типа
            extra = _normalize_text_cols(extra)
            if "mode" not in extra.columns:
                extra["mode"] = "в работе"

            df["well"] = df["well"].astype(str)          # ← та же нормализация у базовых данных
            df = pd.concat([df, extra], ignore_index=True)
            df = df.drop_duplicates(subset=["timestamp", "well"], keep="last")
            df = df.sort_values("timestamp").reset_index(drop=True)
        except Exception as e:
            print(f"[WARN] Ошибка чтения доп.данных: {e}")

    return df


def _load_base_data() -> pd.DataFrame:
    """Читает Excel или возвращает демо-данные."""
    if os.path.exists(EXCEL_PATH):
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=EXCEL_SHEET, parse_dates=["timestamp"])
            df.columns = df.columns.str.strip().str.lower()
            required = {"timestamp","field","well","electricity_fact","electricity_plan",
                        "liquid","oil","ure_fact","ure_plan"}
            missing = required - set(df.columns)
            if missing:
                raise ValueError(f"Отсутствуют колонки: {missing}")
            if "mode" not in df.columns:
                df["mode"] = "в работе"
            df["timestamp"] = pd.to_datetime(df["timestamp"])
            df = _normalize_text_cols(df)
            return df
        except Exception as e:
            print(f"[WARN] Ошибка чтения Excel: {e}\n→ Используются демо-данные.")
    else:
        print(f"[INFO] Файл '{EXCEL_PATH}' не найден → используются демо-данные.")

    return generate_demo_data()


def filter_df(df: pd.DataFrame, field: str, period: str) -> pd.DataFrame:
    if field and field != "ALL":
        df = df[df["field"] == field]
    delta = PERIODS.get(period, timedelta(days=30))
    cutoff = df["timestamp"].max() - delta
    return df[df["timestamp"] >= cutoff]


# ════════════════════════════════════════════════
#  ПОСТРОЕНИЕ ГРАФИКОВ
# ════════════════════════════════════════════════

import math as _math

# Границы зон на шкале (макс = 31.5, последняя зона — «>31.5»)
GAUGE_MAX   = 35
GAUGE_ZONES = [
    (0,    5,    GREEN_DARK,    "0"),
    (5,    10,   GREEN_MID,     "5"),
    (10,   15,   GREEN_LIGHT,   "10"),
    (15,   20,   GREEN_TRANSP,  "15"),
    (20,   25,   YELLOW_TRANSP, "20"),
    (25,   30,   YELLOW_LIGHT,  "25"),
    (30,   35,   YELLOW,        "30"),
]

# Центр дуги в paper-координатах и вертикальное сжатие gauge
_CX, _CY, _SQUEEZE = 0.5, 0.30, 0.85


def _angle(value: float) -> float:
    """180° (лево) → 0° (право) по значению шкалы."""
    v = min(max(value, 0), GAUGE_MAX)
    return _math.radians(180 - (v / GAUGE_MAX) * 180)


def _arc_xy(value: float, r: float) -> tuple:
    """Paper-координаты точки на дуге."""
    a = _angle(value)
    return (
        round(_CX + r * _math.cos(a), 4),
        round(_CY + r * _math.sin(a) * _SQUEEZE, 4),
    )


def _needle_trace(value: float) -> go.Scatter:
    """Стрелка-треугольник; основание опущено ниже центра."""
    angle = _angle(value)
    L  = 0.34   # длина стрелки
    hw = 0.018  # полуширина основания

    tip_x = _CX + L * _math.cos(angle)
    tip_y = _CY + L * _math.sin(angle) * _SQUEEZE

    # основание — ниже центра
    base_y_offset = -0.17
    perp = angle + _math.pi / 2
    bx1 = _CX + hw * _math.cos(perp)
    by1 = _CY + base_y_offset + hw * _math.sin(perp) * _SQUEEZE
    bx2 = _CX - hw * _math.cos(perp)
    by2 = _CY + base_y_offset - hw * _math.sin(perp) * _SQUEEZE

    return go.Scatter(
        x=[bx1, tip_x, bx2, bx1],
        y=[by1, tip_y, by2, by1],
        mode="lines", fill="toself",
        fillcolor="#37474f",
        line=dict(color="#37474f", width=1),
        xaxis="x", yaxis="y",
        showlegend=False, hoverinfo="skip",
    )


def make_gauge(ure_val: float, df: pd.DataFrame = None):
    """
    Спидометр:
    • шкала 0–31.5, последняя метка «>31.5»
    • метки границ — снизу дуги (ticktext через ось)
    • количество скважин — сверху дуги (аннотации)
    • стрелка с опущенным основанием
    • белый фон, без серой заливки
    """
    # ── подсчёт скважин по зонам ─────────────────
    zone_counts = [0] * len(GAUGE_ZONES)
    if df is not None and "ure_fact" in df.columns and "well" in df.columns:
        well_ure = df.groupby("well")["ure_fact"].mean()
        for w_val in well_ure:
            placed = False
            for i, (lo, hi, *_) in enumerate(GAUGE_ZONES):
                if lo <= w_val < hi:
                    zone_counts[i] += 1
                    placed = True
                    break
            if not placed:          # ≥ 31.5 → последняя зона
                zone_counts[-1] += 1

    # ── шаги дуги ────────────────────────────────
    steps = [{"range": [lo, hi], "color": clr}
             for lo, hi, clr, *_ in GAUGE_ZONES]

    # Метки границ на оси (снизу дуги — стандартное поведение gauge)
    tick_vals = [z[0] for z in GAUGE_ZONES] + [GAUGE_MAX]
    tick_text = [z[3] for z in GAUGE_ZONES] + [f">{int(GAUGE_MAX)}"]

    gauge_trace = go.Indicator(
        mode="gauge",
        value=min(ure_val, GAUGE_MAX),
        gauge={
            "axis": {
                "range":    [0, GAUGE_MAX],
                "tickvals": tick_vals,
                "ticktext": tick_text,
                "tickfont": {"size": 10, "color": "#555"},
                "tickwidth": 1,
                "tickcolor": "#aaa",
                "ticklen": 6,
            },
            "bar":       {"color": "rgba(0,0,0,0)", "thickness": 0},
            "bgcolor":   "rgba(0,0,0,0)",   # убираем серый фон
            "borderwidth": 0,
            "steps":     steps,
        },
        domain={"x": [0, 1], "y": [0, 0.92]},
    )

    fig = go.Figure(gauge_trace)

    # ── стрелка ───────────────────────────────────
    fig.add_trace(_needle_trace(ure_val))

    # ── точка в центре ────────────────────────────
    fig.add_trace(go.Scatter(
        x=[_CX], y=[_CY - 0.17], mode="markers",
        marker=dict(color="#37474f", size=9),
        xaxis="x", yaxis="y",
        showlegend=False, hoverinfo="skip",
    ))

    # ── значение под центром ──────────────────────
    fig.add_annotation(
        x=0.5, y=-0.02,
        text=f"<b>{ure_val:.2f} кВт·ч/м³</b>",
        showarrow=False,
        font=dict(size=13, color=TEXT_MUTED),
        xref="paper", yref="paper",
        xanchor="center",
    )

    # ── подписи кол-ва скважин СВЕРХУ каждой зоны ─
    annotations = list(fig.layout.annotations)  # уже содержит значение
    #for i, (lo, hi, clr, _tick) in enumerate(GAUGE_ZONES):
    #    mid = (lo + hi) / 2
    #    # r=0.62 — над дугой (дуга примерно r=0.50 в paper-единицах)
    #    rx, ry = _arc_xy(mid, r=0.64)
    #    annotations.append(dict(
    #        x=rx, y=ry,
    #        text=f"<b>{zone_counts[i]}<br>скв</b>",
    #        showarrow=False,
    #        font=dict(size=8, color="crimson"),
    #        xref="paper", yref="paper",
    #        xanchor="center", yanchor="bottom",
    #        align="center",
    #    ))

    fig.update_layout(
        height=240,
        margin=dict(l=50, r=50, t=5, b=30),
        paper_bgcolor=CARD_BG,
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Arial"},
        annotations=annotations,
        xaxis=dict(visible=False, range=[0, 1]),
        yaxis=dict(visible=False, range=[0, 1]),
    )
    return fig


def make_top_best(df: pd.DataFrame):
    df = df[df["electricity_fact"] > 0]
    agg = df.groupby("well")["ure_fact"].mean().nsmallest(10).sort_values(ascending=True)
    colors = [GREEN_DARK if v<7.5 else GREEN_MID if v<15 else GREEN_LIGHT if v<22.5 else GREEN_TRANSP for v in agg.values]
    fig = go.Figure(go.Bar(
        x=agg.values, y=agg.index.astype(str), orientation="h",
        marker_color=colors,
        text=[f"{v:.1f}" for v in agg.values], textposition="outside",
        textfont=dict(size=11, color=GREY_DARK),
        hovertemplate="%{y}: %{x:.1f} кВт·ч/м³<extra></extra>"
    ))
    fig.update_layout(
        height=240, margin=dict(l=50,r=60,t=10,b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(range=[0, agg.max()*1.3], showgrid=False, zeroline=False, visible=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_top_worst(df: pd.DataFrame):
    df = df[df["electricity_fact"] > 0]
    agg = df.groupby("well")["ure_fact"].mean().nlargest(10).sort_values(ascending=False)
    colors = [ORANGE if v>37.5 else ORANGE_MID if v>30 else ORANGE_LIGHT if v>22.5 else ORANGE_TRANSP for v in agg.values]
    fig = go.Figure(go.Bar(
        x=agg.values, y=agg.index.astype(str), orientation="h",
        marker_color=colors,
        text=[f"{v:.1f}" for v in agg.values], textposition="outside",
        textfont=dict(size=11, color=GREY_DARK),
        hovertemplate="%{y}: %{x:.1f} кВт·ч/м³<extra></extra>"
    ))
    fig.update_layout(
        height=240, margin=dict(l=50,r=60,t=10,b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(range=[0, agg.max()*1.3], showgrid=False, zeroline=False, visible=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_consumption(df: pd.DataFrame):
    df["well"] = df["well"].astype(str)
    agg = df.groupby("well").agg(
        fact=("electricity_fact","sum"),
        plan=("electricity_plan","sum")
    ).reset_index().sort_values("well")
    agg["fact_k"] = (agg["fact"] / 1000).round(1)
    agg["plan_k"] = (agg["plan"] / 1000).round(1)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Фактическое потребление", x=agg["well"], y=agg["fact_k"],
        marker_color=GREY_DARK,
        text=agg["fact_k"], textposition="outside", textfont=dict(size=11),
        hovertemplate="%{x}: %{y} тыс.кВт·ч<extra></extra>"
    ))
    fig.add_trace(go.Bar(
        name="Расчётное потребление", x=agg["well"], y=agg["plan_k"],
        marker_color=GREY_MID,
        text=agg["plan_k"], textposition="outside", textfont=dict(size=11),
        hovertemplate="%{x}: %{y} тыс.кВт·ч<extra></extra>"
    ))
    ymax = max(agg["fact_k"].max(), agg["plan_k"].max()) * 1.25 if len(agg) else 100
    fig.update_layout(
        barmode="group", bargap=0.15, bargroupgap=0.02,
        height=250, margin=dict(l=40,r=150,t=30,b=40),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(range=[0, ymax], gridcolor="#eeeeee", title="тыс.кВт·ч", title_font=dict(size=11)),
        xaxis=dict(tickfont=dict(size=11, weight='bold')),
        legend=dict(orientation="h", y=1.15, x=0, font=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_trend(df: pd.DataFrame):
    """Тренд УРЭ факт/расчёт по времени."""
    agg = df.groupby("timestamp").agg(ure_fact=("ure_fact", "mean"), ure_plan=("ure_plan", "mean")).reset_index()
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=agg["timestamp"], y=agg["ure_plan"], mode="lines", name="УРЭ расч",
        line=dict(color=GREY_LIGHT, width=2, dash="dot"),
        hovertemplate="УРЭ расч: %{y:.2f} кВт·ч/м³<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=agg["timestamp"], y=agg["ure_fact"], mode="lines+markers", name="УРЭ факт",
        line=dict(color=ORANGE, width=2), marker=dict(size=4), fill="tonexty",
        fillcolor=with_alpha(ORANGE, 0.08),
        hovertemplate="УРЭ факт: %{y:.2f} кВт·ч/м³<extra></extra>",
    ))
    fig.update_layout(
        height=240, margin=dict(l=50, r=20, t=25, b=40),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(gridcolor="#eeeeee", title="кВт·ч/м³", title_font=dict(size=10)),
        xaxis=dict(gridcolor="#eeeeee", showspikes=True, spikemode="across",
                   spikesnap="cursor", spikedash="solid", spikecolor=TEXT_MUTED, spikethickness=1),
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor=with_alpha(BG, 0.8),
            bordercolor=GREY_LIGHT,
            font=dict(family="Arial", size=11, color=TEXT_DARK),
        ),
        legend=dict(orientation="h", y=1.15, x=0, font=dict(size=10)),
        font={"family": "Arial", "color": TEXT_DARK},
    )
    return fig


def make_consumption_trend(df: pd.DataFrame):
    """Суточная динамика потребления электроэнергии факт/план,
    с заливкой превышения (оранжевым) и экономии (зелёным)."""
    d = df.copy()
    # d["date"] = d["timestamp"].dt.date
    agg = d.groupby("timestamp").agg(
        fact=("electricity_fact", "sum"), plan=("electricity_plan", "sum")
    ).reset_index()
    agg["fact_k"] = agg["fact"] / 1000
    agg["plan_k"] = agg["plan"] / 1000

    agg["overage_k"] = np.where(agg["fact_k"] > agg["plan_k"], agg["fact_k"], agg["plan_k"])
    agg["underage_k"] = np.where(agg["fact_k"] < agg["plan_k"], agg["fact_k"], agg["plan_k"])

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=agg["timestamp"], y=agg["plan_k"], name="План", mode="lines",
        line=dict(color=GREY_LIGHT, width=2, dash="dot"),
        hovertemplate="План: %{y:.2f} тыс.кВт·ч<extra></extra>",
    ))

    fig.add_trace(go.Scatter(
        x=agg["timestamp"], y=agg["fact_k"], name="Факт", mode="lines+markers",
        line=dict(color=GREY_DARK, width=2), marker=dict(size=4),
        fill="tonexty", fillcolor=with_alpha(ORANGE, 0.08),
        hovertemplate="Факт: %{y:.2f} тыс.кВт·ч<extra></extra>",
    ))

    fig.update_layout(
        height=240, margin=dict(l=50, r=20, t=25, b=40),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(gridcolor="#eeeeee", title="тыс.кВт·ч", title_font=dict(size=10)),
        xaxis=dict(gridcolor="#eeeeee", showspikes=True, spikemode="across",
                   spikesnap="cursor", spikedash="solid", spikecolor=TEXT_MUTED, spikethickness=1),
        hovermode="x unified",
        hoverlabel=dict(
            bgcolor=with_alpha(BG, 0.8),
            bordercolor=GREY_LIGHT,
            font=dict(family="Arial", size=11, color=TEXT_DARK),
        ),
        legend=dict(orientation="h", y=1.15, x=0, font=dict(size=10)),
        font={"family": "Arial", "color": TEXT_DARK},
    )

    return fig




def make_overconsumption_bar(df: pd.DataFrame, top_n: int = 8):
    """ТОП скважин по суммарному перерасходу электроэнергии (факт − план)."""
    agg = df.groupby("well").agg(
        fact=("electricity_fact", "sum"), plan=("electricity_plan", "sum"),
    ).reset_index()
    agg["excess"] = agg["fact"] - agg["plan"]
    agg = agg[agg["excess"] > 0].sort_values("excess", ascending=False).head(top_n)
    if agg.empty:
        fig = go.Figure()
        fig.update_layout(
            annotations=[dict(text="Перерасход не обнаружен", xref="paper", yref="paper",
                               x=0.5, y=0.5, showarrow=False, font=dict(size=13, color=TEXT_MUTED))],
            paper_bgcolor=CARD_BG, height=240,
        )
        return fig
    agg = agg.sort_values("excess", ascending=True)
    fig = go.Figure(go.Bar(
        x=agg["excess"] / 1000, y=agg["well"].astype(str), orientation="h",
        marker_color=with_alpha(ORANGE, 0.75),
        text=[f"{v/1000:.1f}" for v in agg["excess"]], textposition="outside",
        textfont=dict(size=11, color=TEXT_DARK),
        hovertemplate="Скв. %{y}: +%{x:.1f} тыс.кВт·ч сверх плана<extra></extra>",
    ))
    fig.update_layout(
        height=240, margin=dict(l=50, r=50, t=15, b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(title="тыс.кВт·ч сверх плана", title_font=dict(size=10), showgrid=False),
        yaxis=dict(tickfont=dict(size=11)),
        font={"family": "Arial", "color": TEXT_DARK},
    )
    return fig


def build_quick_summary(df: pd.DataFrame):
    """Карточка «Быстрая сводка» — доли скважин по статусу УРЭ (как на референсе)."""
    well_agg = df.groupby("well").agg(
        ure_fact=("ure_fact", "mean"), ure_plan=("ure_plan", "mean"),
    ).reset_index()
    well_agg["dev"] = np.where(well_agg["ure_plan"] != 0,
                                (well_agg["ure_fact"] - well_agg["ure_plan"]) / well_agg["ure_plan"] * 100, 0)
    total = len(well_agg)
    if total == 0:
        return html.Div()

    norm = int((well_agg["dev"] <= 5).sum())
    warn = int(((well_agg["dev"] > 5) & (well_agg["dev"] <= 15)).sum())
    crit = int((well_agg["dev"] > 15).sum())

    rows = [
        ("В норме",              norm, GREEN_DARK),
        ("Незначит. отклонения", warn, YELLOW),
        ("Критические отклонения", crit, ORANGE),
    ]

    items = []
    for label, count, color in rows:
        pct = round(count / total * 100) if total else 0
        items.append(html.Div([
            html.Div([
                html.Span(style={
                    "display": "inline-block", "width": "8px", "height": "8px", "borderRadius": "50%",
                    "background": color, "marginRight": "8px",
                }),
                html.Span(label, style={"fontSize": "12.5px", "color": TEXT_DARK, "flex": "1"}),
                html.Span(str(count), style={"fontSize": "12.5px", "fontWeight": "700", "color": TEXT_DARK, "marginRight": "8px"}),
                html.Span(f"{pct}%", style={"fontSize": "11px", "color": TEXT_MUTED, "width": "34px", "textAlign": "right"}),
            ], style={"display": "flex", "alignItems": "center", "marginBottom": "4px"}),
            html.Div(html.Div(style={
                "width": f"{pct}%", "height": "6px", "borderRadius": "4px", "background": color,
            }), style={"background": "#eef0f4", "borderRadius": "4px", "height": "6px", "marginBottom": "12px"}),
        ]))

    return html.Div([
        html.Div([
            html.Span("Быстрая сводка", style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
        ], style={"marginBottom": "12px"}),
        html.Div(items),
    ])


def generate_insights(df: pd.DataFrame):
    """Автоматические текстовые выводы по фонду скважин за выбранный период."""
    if df.empty:
        return []

    well_agg = df.groupby(["field", "well"]).agg(
        ure_fact=("ure_fact", "mean"), ure_plan=("ure_plan", "mean"),
        fact=("electricity_fact", "sum"), plan=("electricity_plan", "sum"),
    ).reset_index()
    well_agg["dev"] = np.where(well_agg["ure_plan"] != 0,
                                (well_agg["ure_fact"] - well_agg["ure_plan"]) / well_agg["ure_plan"] * 100, 0)

    total_wells = len(well_agg)
    critical = well_agg[well_agg["dev"] > 15]
    excess_total = (well_agg["fact"] - well_agg["plan"]).clip(lower=0).sum()

    field_dev = well_agg.groupby("field")["dev"].mean().sort_values(ascending=False)
    worst_field = field_dev.index[0] if len(field_dev) else "—"
    best_field = field_dev.index[-1] if len(field_dev) else "—"

    worst_well = well_agg.sort_values("dev", ascending=False).iloc[0] if total_wells else None

    insights = []
    if critical.shape[0]:
        insights.append({
            "icon": "trend", "color": ORANGE,
            "title": f"Высокое отклонение УРЭ у {critical.shape[0]} скважин",
            "text": f"{critical.shape[0]} из {total_wells} скважин "
                    f"({critical.shape[0]/total_wells*100:.0f}%) превышают +15% от плана.",
        })
    insights.append({
        "icon": "bolt", "color": GREEN_DARK,
        "title": f"Перерасход энергопотребления",
        "text": f"Суммарный перерасход электроэнергии за период составляет: "
                f"{excess_total/1000:,.0f} тыс.кВт·ч.".replace(",", " "),
    })
    if worst_field != best_field:
        insights.append({
            "icon": "pin", "color": YELLOW,
            "title": f"Месторождение «{worst_field}» — зона риска",
            "text": f"Среднее отклонение УРЭ {field_dev.iloc[0]:.1f}%, тогда как у «{best_field}» "
                    f"всего {field_dev.iloc[-1]:.1f}%.",
        })
    if worst_well is not None:
        insights.append({
            "icon": "gauge", "color": ORANGE,
            "title": f"Скважина {worst_well['well']} — максимальное отклонение",
            "text": f"{worst_well['field']}: {worst_well['dev']:+.1f}% от расчётного УРЭ. "
                    f"Рекомендуется рассмотреть мероприятие по снижению УРЭ на карточке скважины.",
        })

    return insights

# def build_rating_table(df: pd.DataFrame):
#     """Полный рейтинг скважин фонда — сортируемая таблица."""
#     from dash import dash_table
#
#     agg = df.groupby(["field", "well"]).agg(
#         electricity_fact=("electricity_fact", "sum"),
#         electricity_plan=("electricity_plan", "sum"),
#         liquid=("liquid", "sum"),
#         oil=("oil", "sum"),
#         ure_fact=("ure_fact", "mean"),
#         ure_plan=("ure_plan", "mean"),
#     ).reset_index()
#     agg["dev_pct"] = np.where(agg["ure_plan"] != 0,
#                                (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100, 0).round(1)
#     agg = agg.sort_values("dev_pct", ascending=False).round(2)
#     agg = agg.rename(columns={
#         "field": "Месторождение", "well": "Скважина",
#         "electricity_fact": "ЭЭ факт, кВт·ч", "electricity_plan": "ЭЭ расч, кВт·ч",
#         "liquid": "Жидкость, м³", "oil": "Нефть, т",
#         "ure_fact": "УРЭ факт", "ure_plan": "УРЭ расч", "dev_pct": "Откл. УРЭ, %",
#     })
#
#     return dash_table.DataTable(
#         data=agg.to_dict("records"),
#         columns=[{"name": c, "id": c} for c in agg.columns],
#         sort_action="native", filter_action="native", page_size=20,
#         style_table={"overflowX": "auto"},
#         style_header={
#             "backgroundColor": "#f7f8fb", "fontWeight": "700", "color": TEXT_DARK,
#             "border": "none", "borderBottom": f"1px solid {BORDER_SOFT}", "fontSize": "12px",
#         },
#         style_cell={
#             "fontFamily": "Arial", "fontSize": "12px", "color": TEXT_DARK,
#             "padding": "8px 14px", "border": "none", "borderBottom": f"1px solid {BORDER_SOFT}",
#             "backgroundColor": CARD_BG,
#         },
#         style_data_conditional=[
#             {"if": {"filter_query": "{Откл. УРЭ, %} > 15", "column_id": "Откл. УРЭ, %"},
#              "color": ORANGE, "fontWeight": "700"},
#             {"if": {"filter_query": "{Откл. УРЭ, %} <= 5", "column_id": "Откл. УРЭ, %"},
#              "color": GREEN_DARK, "fontWeight": "700"},
#         ],
#     )

# def build_rating_table(df: pd.DataFrame):
#     """Полный рейтинг скважин фонда — сортируемая таблица (сортировка на стороне Python)."""
#     from dash import dash_table
#
#     df = df.drop_duplicates(subset=["timestamp", "well"], keep="last")
#
#     agg = df.groupby(["field", "well"]).agg(
#         electricity_fact=("electricity_fact", "sum"),
#         electricity_plan=("electricity_plan", "sum"),
#         liquid=("liquid", "sum"),
#         oil=("oil", "sum"),
#         ure_fact=("ure_fact", "mean"),
#         ure_plan=("ure_plan", "mean"),
#     ).reset_index()
#
#     agg["dev_pct"] = np.where(agg["ure_plan"] != 0,
#                                (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100, 0).round(1)
#     agg = agg.sort_values("dev_pct", ascending=False, kind="mergesort").round(2)
#
#     agg = agg.rename(columns={
#         "field": "Месторождение", "well": "Скважина",
#         "electricity_fact": "ЭЭ факт, кВт·ч", "electricity_plan": "ЭЭ расч, кВт·ч",
#         "liquid": "Жидкость, м³", "oil": "Нефть, т",
#         "ure_fact": "УРЭ факт", "ure_plan": "УРЭ расч", "dev_pct": "Откл. УРЭ, %",
#     })
#
#     TEXT_COLS = {"Месторождение", "Скважина"}
#
#     columns = [
#         {"name": c, "id": c, "type": "text" if c in TEXT_COLS else "numeric"}
#         for c in agg.columns
#     ]
#
#     return dash_table.DataTable(
#         id="rating-table",
#         data=agg.to_dict("records"),
#         columns=columns,
#         sort_action="custom",
#         sort_mode="single",
#         filter_action="native", page_size=20,
#         sort_by=[{"column_id": "Откл. УРЭ, %", "direction": "desc"}],
#         style_table={"overflowX": "auto", "height": "780px", "overflowY": "auto"},
#         style_header={
#             "backgroundColor": "#f7f8fb", "fontWeight": "700", "color": TEXT_DARK,
#             "border": "none", "borderBottom": f"1px solid {BORDER_SOFT}", "fontSize": "12px",
#         },
#         style_cell={
#             "fontFamily": "Arial", "fontSize": "12px", "color": TEXT_DARK,
#             "padding": "8px 10px", "border": "none", "borderBottom": f"1px solid {BORDER_SOFT}",
#             "backgroundColor": CARD_BG,
#         },
#         style_data_conditional=[
#             {"if": {"filter_query": "{Откл. УРЭ, %} > 15", "column_id": "Откл. УРЭ, %"},
#              "color": ORANGE, "fontWeight": "700"},
#             {"if": {"filter_query": "{Откл. УРЭ, %} <= 5", "column_id": "Откл. УРЭ, %"},
#              "color": GREEN_DARK, "fontWeight": "700"},
#         ],
#     )


# ════════════════════════════════════════════════
#  ТАБЛИЦА / РЕЙТИНГ СКВАЖИНЫ
# ════════════════════════════════════════════════

def build_rating_table(df: pd.DataFrame):
    """Рейтинг скважин фонда: агрегация за период + сортируемая/фильтруемая таблица."""

    df = df.drop_duplicates(subset=["timestamp", "well"], keep="last")

    agg = df.groupby(["field", "well"]).agg(
        electricity_fact=("electricity_fact", "sum"),
        electricity_plan=("electricity_plan", "sum"),
        liquid=("liquid", "sum"),
        oil=("oil", "sum"),
        ure_fact=("ure_fact", "mean"),
        ure_plan=("ure_plan", "mean"),
    ).reset_index()

    agg["dev_pct"] = np.where(
        agg["ure_plan"] != 0,
        (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100,
        0,
    )
    agg = agg.round(1).sort_values("dev_pct", ascending=False, kind="mergesort")

    agg = agg.rename(columns={
        "field": "Месторождение",
        "well": "Скважина",
        "electricity_fact": "ЭЭ факт, кВт·ч",
        "electricity_plan": "ЭЭ расч, кВт·ч",
        "liquid": "Жидкость, м³",
        "oil": "Нефть, т",
        "ure_fact": "УРЭ факт",
        "ure_plan": "УРЭ расч",
        "dev_pct": "Откл. УРЭ, %",
    })

    TEXT_COLUMNS = {"Месторождение", "Скважина"}
    columns = [
        {"name": c, "id": c, "type": "text" if c in TEXT_COLUMNS else "numeric"}
        for c in agg.columns
    ]

    return dash_table.DataTable(
        id="rating-table",
        data=agg.to_dict("records"),
        columns=columns,
        sort_action="custom",
        filter_action="native",
        page_size=20,
        style_table={"overflowX": "auto", "height": "780px"},
        style_header={
            "backgroundColor": "#f7f8fb",
            "fontWeight": "700",
            "border": "none",
            "borderBottom": "1px solid #e0e0e0",
            "fontSize": "12px",
            "color": TEXT_DARK,
        },
        style_cell={
            "fontFamily": "Arial",
            "fontSize": "12px",
            "color": TEXT_DARK,
            "padding": "8px 10px",
            "border": "none",
            "borderBottom": "1px solid #e0e0e0",
        },
        style_data_conditional=[
            {"if": {"filter_query": "{Откл. УРЭ, %} > 15", "column_id": "Откл. УРЭ, %"},
             "color": ORANGE, "fontWeight": "700"},
            {"if": {"filter_query": "{Откл. УРЭ, %} <= 5", "column_id": "Откл. УРЭ, %"},
             "color": GREEN_DARK, "fontWeight": "700"},
        ],
    )



# ════════════════════════════════════════════════
#  СТАТУС / ОТКЛОНЕНИЕ СКВАЖИНЫ
# ════════════════════════════════════════════════

def well_status(dev_pct: float):
    """Возвращает (цвет, подпись) по отклонению УРЭ факт от расчёта."""
    if dev_pct <= 5:
        return GREEN_DARK, "норма"
    elif dev_pct <= 15:
        return YELLOW, "внимание"
    return ORANGE, "критично"


# ════════════════════════════════════════════════
#  ЛЕВАЯ ПАНЕЛЬ — МЕСТОРОЖДЕНИЯ → СКВАЖИНЫ (поиск + сворачивание)
# ════════════════════════════════════════════════

def _well_card(row):
    no_data = row["no_data"]
    if no_data:
        color, _label = GREY_LIGHT, "нет данных"
    else:
        color, _label = well_status(row["dev"])
    return html.Div([
        html.Div([
            html.Span(style={
                "display": "inline-block", "width": "8px", "height": "8px",
                "borderRadius": "50%", "background": color, "marginRight": "7px",
                "flexShrink": "0",
            }),
            html.Span(f"Скв. {row['well']}", style={"fontWeight": "700", "fontSize": "13px", "color": TEXT_DARK}),
            html.Span("—" if no_data else f"{row['dev']:+.1f}%", style={
                "marginLeft": "auto", "fontSize": "11px", "fontWeight": "700", "color": color
            }),
        ], style={"display": "flex", "alignItems": "center"}),
        html.Div(f"УРЭ {row['ure_fact']:.1f} кВт·ч/м³",
                 style={"fontSize": "11px", "color": TEXT_MUTED, "marginTop": "2px"}),
    ],
        id={"type": "well-card", "index": str(row["well"])}, n_clicks=0,
        style={
            "background": with_alpha(color, 0.06), "border": f"1px solid {with_alpha(color, 0.18)}",
            "borderRadius": "10px", "padding": "8px 10px", "marginBottom": "6px", "cursor": "pointer",
        }
    )


def build_sidebar_wells(df: pd.DataFrame, search_term: str = ""):
    if df.empty:
        return html.Div("Нет данных", style={"fontSize": "12px", "color": TEXT_MUTED, "padding": "8px"})

    latest = df.sort_values("timestamp").groupby("well").last().reset_index()[["well", "field"]]
    agg = df.groupby("well").agg(
        ure_fact=("ure_fact", "mean"),
        ure_plan=("ure_plan", "mean"),
        liquid=("liquid", "sum"),
        electricity=("electricity_fact", "sum"),  # ← добавить для проверки активности
    ).reset_index()
    agg["dev"] = np.where(agg["ure_plan"] != 0,
                          (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100, 0)
    agg["no_data"] = (agg["electricity"] == 0) & (agg["liquid"] == 0)

    agg = agg.merge(latest, on="well", how="left")

    term = (search_term or "").strip().lower()

    groups = []
    for field_name in sorted(agg["field"].dropna().unique()):
        g = agg[agg["field"] == field_name].sort_values("dev", ascending=False)
        if term:
            g = g[g["well"].astype(str).str.lower().str.contains(term)]
        if g.empty:
            continue

        well_cards = [_well_card(row) for _, row in g.iterrows()]

        groups.append(html.Details([
            html.Summary([
                svg_icon("pin", color=BLUE, size=14),
                html.Span(field_name, style={
                    "marginLeft": "6px", "fontWeight": "700", "fontSize": "13px", "color": TEXT_DARK,
                }),
                html.Span(str(len(g)), style={
                    "marginLeft": "auto", "fontSize": "11px", "color": TEXT_MUTED,
                    "background": "#f0f1f5", "padding": "1px 8px", "borderRadius": "9px",
                }),
            ], style={
                "display": "flex", "alignItems": "center", "cursor": "pointer",
                "padding": "6px 4px", "gap": "2px",
            }),
            html.Div(well_cards, style={"marginTop": "4px", "marginLeft": "2px"}),
        ], open=True, style={
            "marginBottom": "6px", "borderBottom": f"1px solid {BORDER_SOFT}", "paddingBottom": "6px",
        }))

    if not groups:
        return html.Div("Скважины не найдены", style={"fontSize": "12px", "color": TEXT_MUTED, "padding": "8px"})
    return html.Div(groups)


# ════════════════════════════════════════════════
#  ДЕТАЛЬНАЯ КАРТОЧКА СКВАЖИНЫ (окно по клику)
# ════════════════════════════════════════════════

def build_well_trend_fig(df_period: pd.DataFrame, well: str):
    d = df_period[df_period["well"].astype(str) == str(well)].sort_values("timestamp")
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=d["timestamp"], y=d["ure_plan"], name="УРЭ расч",
        line=dict(color=GREY_LIGHT, width=2, dash="dot"),
        hovertemplate="%{x|%d.%m %H:%M}: %{y:.2f}<extra>УРЭ расч</extra>"
    ))
    fig.add_trace(go.Scatter(
        x=d["timestamp"], y=d["ure_fact"], name="УРЭ факт",
        line=dict(color=ORANGE, width=2),
        hovertemplate="%{x|%d.%m %H:%M}: %{y:.2f}<extra>УРЭ факт</extra>"
    ))
    fig.add_trace(go.Scatter(
        x=d["timestamp"], y=d["liquid"], name="Дебит жидкости",
        yaxis="y2", line=dict(color=GREEN_DARK, width=2),
        hovertemplate="%{x|%d.%m %H:%M}: %{y:.0f} м³<extra>Дебит жидкости</extra>"
    ))
    fig.update_layout(
        height=230, margin=dict(l=45, r=45, t=30, b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(title="кВт·ч/м³", title_font=dict(size=10), gridcolor="#eeeeee"),
        yaxis2=dict(title="м³", title_font=dict(size=10), overlaying="y", side="right", showgrid=False),
        legend=dict(orientation="h", y=1.22, x=0, font=dict(size=10)),
        xaxis=dict(gridcolor="#eeeeee"),
        font={"family": "Arial"},
    )
    return fig


def build_well_compare_bar(df_period: pd.DataFrame, well: str):
    d = df_period[df_period["well"].astype(str) == str(well)]
    field = d["field"].iloc[0] if not d.empty else None
    well_ure = d["ure_fact"].mean() if not d.empty else 0
    field_df = df_period[df_period["field"] == field] if field else df_period
    field_avg = field_df["ure_fact"].mean() if not field_df.empty else 0

    fig = go.Figure(go.Bar(
        x=["Скважина", "Среднее по м/р"], y=[well_ure, field_avg],
        marker_color=[ORANGE if well_ure > field_avg else GREEN_DARK, GREY_LIGHT],
        text=[f"{well_ure:.1f}", f"{field_avg:.1f}"], textposition="outside",
        textfont=dict(size=12, color=GREY_DARK),
        hovertemplate="%{x}: %{y:.2f} кВт·ч/м³<extra></extra>",
    ))
    ymax = max(well_ure, field_avg) * 1.35 if max(well_ure, field_avg) else 1
    fig.update_layout(
        height=230, margin=dict(l=30, r=30, t=30, b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(visible=False, range=[0, ymax]),
        xaxis=dict(tickfont=dict(size=12)),
        showlegend=False, font={"family": "Arial"},
    )
    return fig


def load_meropriyatiya() -> pd.DataFrame:
    """Читает Excel с мероприятиями по снижению УРЭ (если загружен)."""
    if os.path.exists(MEROPRIYATIYA_PATH):
        try:
            df = pd.read_excel(MEROPRIYATIYA_PATH)
            df.columns = [str(c).strip() for c in df.columns]
            if "Скважина" in df.columns:
                df["Скважина"] = df["Скважина"].astype(str)
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=MEROPRIYATIYA_COLUMNS)


def build_meropriyatiya_section(well: str):
    """Раздел «Мероприятия по снижению УРЭ» для конкретной скважины — отдельный лист карточки."""
    df = load_meropriyatiya()
    file_exists = os.path.exists(MEROPRIYATIYA_PATH)

    link_row = html.Div([
        svg_icon("link", color=GREEN_DARK, size=14),
        html.A("Открыть полный Excel-файл с мероприятиями", href="/meropriyatiya.xlsx",
               target="_blank", style={"marginLeft": "6px", "fontSize": "12px", "color": GREEN_DARK,
                                        "fontWeight": "600", "textDecoration": "none"}),
    ], style={"display": "flex", "alignItems": "center", "marginBottom": "10px"}) if file_exists else None

    if not file_exists or df.empty:
        body = html.Div([
            html.Div(
                html.Div("Раздел в разработке", style={
                    "fontSize": "13px", "color": TEXT_MUTED, "fontWeight": "600",
                }),
                style={"display": "flex", "alignItems": "center", "justifyContent": "center", "height": "350px"},
            ),

            html.Div(
                # "Загрузите Excel-файл с колонками «Скважина, Мероприятие, Рекомендации, "
                # "Возможная экономия электроэнергии, кВт» на вкладке «Рейтинг скважин» → блок "
                # "«Мероприятия по снижению УРЭ», либо поместите файл вручную по пути "
                # f"{MEROPRIYATIYA_PATH} рядом со скриптом.",
                style={"fontSize": "12px", "color": TEXT_MUTED, "marginTop": "4px", "lineHeight": "1.5"},
            ),
        ], style={"padding": "16px", "background": "#f7f8fb", "borderRadius": "10px"})
        return html.Div([link_row, body] if link_row else [body])

    well_rows = df[df["Скважина"] == str(well)] if "Скважина" in df.columns else df.iloc[0:0]

    if well_rows.empty:
        body = html.Div(f"Для скважины {well} мероприятий не найдено в загруженном файле.",
                         style={"fontSize": "12px", "color": TEXT_MUTED, "padding": "12px"})
        return html.Div([link_row, body] if link_row else [body])

    header_cells = [c for c in MEROPRIYATIYA_COLUMNS if c in well_rows.columns and c != "Ссылка"]
    table_head = html.Thead(html.Tr([
        html.Th(c, style={"textAlign": "left", "fontSize": "11px", "color": TEXT_MUTED,
                           "padding": "6px 8px", "borderBottom": f"1px solid {BORDER_SOFT}"})
        for c in header_cells + (["Ссылка"] if "Ссылка" in well_rows.columns else [])
    ]))
    body_rows = []
    for _, r in well_rows.iterrows():
        cells = [html.Td(str(r.get(c, "")), style={"fontSize": "12px", "color": TEXT_DARK,
                                                      "padding": "6px 8px", "borderBottom": f"1px solid {BORDER_SOFT}"})
                  for c in header_cells]
        if "Ссылка" in well_rows.columns:
            url = str(r.get("Ссылка", "") or "")
            link_cell = html.Td(
                html.A("Открыть →", href=url, target="_blank",
                       style={"fontSize": "12px", "color": GREEN_DARK, "fontWeight": "600", "textDecoration": "none"})
                if url and url.lower() != "nan" else "—",
                style={"padding": "6px 8px", "borderBottom": f"1px solid {BORDER_SOFT}"},
            )
            cells.append(link_cell)
        body_rows.append(html.Tr(cells))

    table = html.Table([table_head, html.Tbody(body_rows)], style={"width": "100%", "borderCollapse": "collapse"})
    return html.Div([link_row, table] if link_row else [table])


def build_well_detail(df_period: pd.DataFrame, well: str):
    d = df_period[df_period["well"].astype(str) == str(well)]
    if d.empty:
        return html.Div("Нет данных по скважине за выбранный период.")

    last = d.sort_values("timestamp").iloc[-1]
    ure_f, ure_p = d["ure_fact"].mean(), d["ure_plan"].mean()

    has_activity = (d["electricity_fact"].sum() > 0) or (d["liquid"].sum() > 0)
    no_data = not has_activity

    dev = (ure_f - ure_p) / ure_p * 100 if ure_p else 0

    if no_data:
        color, label = GREY_LIGHT, "нет данных"
    else:
        color, label = well_status(dev)

    mode = last["mode"] if "mode" in last and pd.notna(last["mode"]) else "—"
    mode_color = MODE_COLORS.get(mode, GREY_LIGHT)

    def fmt(n):
        return f"{n:,.0f}".replace(",", " ")

    header = html.Div([
        html.Div([
            html.Span(f"Скважина {well}", style={"fontSize": "18px", "fontWeight": "700", "color": TEXT_DARK}),
            html.Span(last["field"], style={"fontSize": "12px", "color": TEXT_MUTED, "marginLeft": "10px"}),
            html.Span(label, style={
                "marginLeft": "10px", "fontSize": "11px", "fontWeight": "700", "color": "white",
                "background": color, "padding": "2px 10px", "borderRadius": "10px",
            }),
            html.Span(mode, style={
                "marginLeft": "10px", "fontSize": "11px", "fontWeight": "700", "color": "white",
                "background": mode_color, "padding": "2px 10px", "borderRadius": "10px",
                "whiteSpace": "nowrap",
            }),

        ]),
        html.Button("✕", id="btn-close-detail", n_clicks=0, style={
            "border": "none", "background": "none", "fontSize": "20px",
            "cursor": "pointer", "color": TEXT_MUTED, "lineHeight": "1",
        }),
    ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "marginBottom": "14px"})

    kpis = html.Div([
        kpi_card("bolt", "ЭЭ факт, кВт·ч", fmt(d["electricity_fact"].sum()), color=PURPLE),
        kpi_card("plug", "ЭЭ расч, кВт·ч", fmt(d["electricity_plan"].sum()), color=BLUE),
        kpi_card("droplet", "Жидкость, м³", fmt(d["liquid"].sum()), color=TEAL),
        kpi_card("barrel", "Нефть, т", fmt(d["oil"].sum()), color=YELLOW),
        kpi_card("gauge", "УРЭ факт, кВт·ч/м³", f"{ure_f:.2f}", color=GREEN_DARK),
        kpi_card("trend", "Откл. УРЭ, %", "—" if no_data else f"{dev:+.1f}", accent=True),
    ], style={"display": "flex", "gap": "8px", "flexWrap": "wrap", "marginBottom": "14px"})

    charts = html.Div([
        html.Div([
            html.Div("Динамика УРЭ и дебита жидкости",
                     style={"fontWeight": "700", "fontSize": "13px", "marginBottom": "4px", "color": TEXT_DARK}),
            dcc.Graph(figure=build_well_trend_fig(df_period, well), config={"displayModeBar": False}),
        ], style={"flex": "1.4", "minWidth": "300px"}),
        html.Div([
            html.Div("Скважина vs среднее по месторождению",
                     style={"fontWeight": "700", "fontSize": "13px", "marginBottom": "4px", "color": TEXT_DARK}),
            dcc.Graph(figure=build_well_compare_bar(df_period, well), config={"displayModeBar": False}),
        ], style={"flex": "1", "minWidth": "220px"}),
    ], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"})

    sheets = dcc.Tabs([
        dcc.Tab(label="Обзор", value="overview", style=TAB_STYLE, selected_style=TAB_SELECTED_STYLE,
                children=html.Div([kpis, charts], style={"paddingTop": "12px"})),
        dcc.Tab(label="Мероприятия по снижению УРЭ", value="measures", style=TAB_STYLE, selected_style=TAB_SELECTED_STYLE,
                children=html.Div(build_meropriyatiya_section(well), style={"paddingTop": "12px"})),
    ], value="overview", persistence=True, persistence_type="local", style={"borderBottom": f"1px solid {BORDER_SOFT}"})

    return html.Div([header, sheets])


# ════════════════════════════════════════════════
#  ТЕПЛОВАЯ КАРТА (Treemap: группа → скважина)
# ════════════════════════════════════════════════

HEAT_COLORSCALE = [
    [0.0,  "#22C55E"],
    [0.30, "#D8F3DC"],
    [0.55, "#FFF3BF"],
    [0.78, "#FFD6A5"],
    [1.0,  "#F4ACB7"],
]


def make_heatmap(df: pd.DataFrame, group_by: str = "field"):
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            annotations=[dict(text="Нет данных за выбранный период", xref="paper", yref="paper",
                               x=0.5, y=0.5, showarrow=False, font=dict(size=14, color="#999"))],
            paper_bgcolor=CARD_BG, height=400,
        )
        return fig

    group_col = "field" if group_by != "mode" else "mode"

    agg = df.groupby([group_col, "well"]).agg(
        ure_fact=("ure_fact", "mean"),
        ure_plan=("ure_plan", "mean"),
        electricity=("electricity_fact", "sum"),
        liquid=("liquid", "sum"),
    ).reset_index()
    agg["dev"] = np.where(agg["ure_plan"] != 0,
                           (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100, 0)
    agg["dev_clip"] = agg["dev"].clip(-10, 35)
    agg["size"] = agg["liquid"].clip(lower=1)

    groups = agg.groupby(group_col).agg(size=("size", "sum"), dev=("dev_clip", "mean")).reset_index()

    labels  = list(groups[group_col].astype(str)) + list(agg["well"].astype(str))
    parents = [""] * len(groups) + list(agg[group_col].astype(str))
    values  = list(groups["size"]) + list(agg["size"])
    colors  = list(groups["dev"]) + list(agg["dev_clip"])
    text = [f"<b>{g}</b>" for g in groups[group_col].astype(str)] + [
        f"﻿<b>{w}</b><br>ЭЭ {e:,.0f}<br>УРЭ {uf:.1f}/{up:.1f}".replace(",", " ")
        for w, e, uf, up in zip(agg["well"].astype(str), agg["electricity"], agg["ure_fact"], agg["ure_plan"])
    ]
    hover = [f"{g}<extra></extra>" for g in groups[group_col].astype(str)] + [
        f"Скв. {w}<br>УРЭ факт: {uf:.2f} кВт·ч/м³<br>УРЭ расч: {up:.2f} кВт·ч/м³<br>Откл: {dv:+.1f}%<extra></extra>"
        for w, uf, up, dv in zip(agg["well"].astype(str), agg["ure_fact"], agg["ure_plan"], agg["dev"])
    ]

    def _dev_to_color(dev_val: float) -> str:
        import plotly.colors as pc
        dev_clip = max(-10, min(35, dev_val))
        t = (dev_clip - (-10)) / (35 - (-10))
        return pc.sample_colorscale(HEAT_COLORSCALE, [t])[0]

    tile_colors = [_dev_to_color(d) for d in groups["dev"]] + [_dev_to_color(d) for d in agg["dev_clip"]]

    fig = go.Figure(go.Treemap(
        labels=labels, parents=parents, values=values, branchvalues="total",
        text=text, texttemplate="%{text}", textfont=dict(size=11, color=TEXT_DARK),
        hovertemplate=hover,
        marker=dict(
            colors=tile_colors,
            line=dict(width=2, color="#ffffff"),
            # colorbar здесь больше не задаём — им займётся отдельный trace ниже
        ),
        pathbar=dict(visible=True, textfont=dict(size=11, color=TEXT_DARK)),
        root=dict(color=GREY_TRANSP),
    ))

    # ── Невидимый trace только ради легенды-шкалы ──
    fig.add_trace(go.Scatter(
        x=[None], y=[None], mode="markers",
        marker=dict(
            colorscale=HEAT_COLORSCALE, cmin=-10, cmax=35,
            color=[-10, 35], showscale=True,
            colorbar=dict(title="Откл УРЭ,%", thickness=12, len=0.8, tickfont=dict(color=TEXT_DARK)),
        ),
        showlegend=False, hoverinfo="skip",
    ))

    fig.update_layout(
        height=780,  margin=dict(l=5, r=5, t=30, b=5),
        paper_bgcolor=CARD_BG, font={"family": "Arial", "color": TEXT_DARK},
    )
    return fig


# ════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ КОМПОНЕНТЫ
# ════════════════════════════════════════════════

def kpi_card(icon_name, title, value, accent=False, color=None):
    color = color or (ORANGE if accent else GREEN_DARK)
    badge_bg = with_alpha(color, 0.14)
    return html.Div([
        html.Div(svg_icon(icon_name, color=color, size=18), style={
            "width": "36px", "height": "36px", "borderRadius": "10px",
            "background": badge_bg, "display": "flex", "alignItems": "center",
            "justifyContent": "center", "marginRight": "10px", "flexShrink": "0",
        }),
        html.Div([
            html.Div(title, style={"fontSize": "10px", "color": TEXT_MUTED, "lineHeight": "1.3"}),
            html.Div(value, style={"fontSize": "19px", "fontWeight": "700", "color": TEXT_DARK}),
        ])
    ], style={
        "display": "flex", "alignItems": "center",
        "background": CARD_BG, "border": f"1px solid {BORDER_SOFT}", "borderRadius": "12px",
        "padding": "10px 14px", "flex": "1", "minWidth": "150px",
        "boxShadow": glow(color, 0.10),
    })


def mode_panel(df: pd.DataFrame):
    counts = {m: 0 for m in MODES}
    if "mode" in df.columns:
        latest = df.sort_values("timestamp").groupby("well").last()
        for m, cnt in latest["mode"].value_counts().items():
            if m in counts:
                counts[m] = int(cnt)

    total = sum(counts.values()) or 1
    bar = html.Div([
        html.Div(style={
            "background": MODE_COLORS[m],
            "flex": str(max(counts[m], 0.1)),
            "height":"16px",
            "borderRadius": "3px 0 0 3px" if i==0 else ("0 3px 3px 0" if i==len(MODES)-1 else "0")
        }) for i, m in enumerate(MODES)
    ], style={"display":"flex","gap":"3px","marginBottom":"14px"})

    grid = html.Div([
        html.Div([
            html.Div(str(counts[m]), style={"fontSize":"28px","fontWeight":"700","color":"#37474f","lineHeight":"1"}),
            html.Div([
                html.Span("■ ", style={"color":MODE_COLORS[m],"fontSize":"13px"}),
                html.Span(m, style={"fontSize":"13px","color":"#555"}),
            ])
        ], style={"background":"#f9f9f9","borderRadius":"14px","padding":"10px","textAlign":"center"})
        for m in MODES
    ], style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"8px"})

    return html.Div([bar, grid])


def period_buttons(active="1м"):
    buttons = []
    for p in PERIODS:
        active_style = {
            "background": BLUE, "color": "white",
            "border": f"1px solid {BLUE}", "fontWeight": "700"
        }
        idle_style = {
            "background": CARD_BG, "color": GREY_DARK,
            "border": "1px solid #ddd", "fontWeight": "400"
        }
        style = {
            **( active_style if p == active else idle_style ),
            "padding": "5px 14px", "borderRadius": "5px",
            "cursor": "pointer", "fontSize": "13px",
        }
        buttons.append(html.Button(p, id=f"btn-period-{p}", n_clicks=0, style=style))
    return html.Div(buttons, style={"display":"flex","gap":"6px"})


TABS = [
    ("overview", "Обзор", "home"),
    ("heatmap",  "Тепловая карта", "grid"),
    ("trends",   "Основные рекомендации", "chart-line"),
    ("rating",   "Рейтинг скважин", "list"),
]


def tab_buttons(active="overview"):
    buttons = []
    for val, label, icon in TABS:
        is_active = val == active
        color = "#ffffff" if is_active else TEXT_DARK
        style = {
            "background": BLUE if is_active else CARD_BG,
            "color": color,
            "border": f"1px solid {BLUE if is_active else BORDER_SOFT}",
            "fontWeight": "700" if is_active else "500",
            "padding": "7px 16px", "borderRadius": "10px",
            "cursor": "pointer", "fontSize": "13px",
            "display": "flex", "alignItems": "center", "gap": "7px",
        }
        buttons.append(html.Button([
            svg_icon(icon, color=color, size=15), html.Span(label),
        ], id=f"btn-tab-{val}", n_clicks=0, style=style))
    return html.Div(buttons, style={"display": "flex", "gap": "8px", "flexWrap": "wrap"})


# ════════════════════════════════════════════════
#  APP LAYOUT
# ════════════════════════════════════════════════

app = dash.Dash(
    __name__, title="Анализ энергопотребления скважин",
    suppress_callback_exceptions=True,
)
server = app.server

app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            html, body { font-family: 'Inter', Arial, sans-serif; }
            ::-webkit-scrollbar { width: 8px; height: 8px; }
            ::-webkit-scrollbar-thumb { background: #d7d9e0; border-radius: 8px; }
            details > summary { list-style: none; }
            details > summary::-webkit-details-marker { display: none; }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
"""
app.layout = html.Div([

    # ── Хранилище состояния ──────────────────────
    dcc.Store(id="store-period", data="1м", storage_type="local"),
    dcc.Store(id="store-active-tab", data="overview", storage_type="local"),
    dcc.Store(id="store-consumption-tab", data="period"),
    dcc.Store(id="store-selected-well", data=None),
    dcc.Store(id="dummy-filter-fix"),


    # ── Шапка ────────────────────────────────────
    html.Div([
        html.Img(
        src="/assets/logo.png",  # поместите logo.png рядом с dashboard.py в папке assets/
        style={
            "maxWidth": "100%",
            "height": "40px",
            "marginRight": "12px",
        }
    ),
        #html.Div("Анализ энергопотребления скважин",
        #         style={"fontSize":"17px","fontWeight":"700","color":GREEN_DARK}),
        html.Div([
            # Фильтр месторождений
            dcc.Dropdown(
                id="dropdown-field",
                options=[],          # заполняется callback'ом
                placeholder="Все месторождения",
                clearable=True,
                style={"width":"220px","fontSize":"13px"},
                persistence=True, persistence_type="local",
            ),
            # Периоды
            html.Div(id="period-buttons-container", children=period_buttons("1м")),
            # Кнопка обновления
            html.Button(
                "Обновить данные",
                id="btn-refresh", n_clicks=0,
                style={
                    "background": GREEN_DARK, "color":"white",
                    "border":"none","borderRadius":"6px",
                    "padding":"7px 16px","cursor":"pointer","fontSize":"13px",
                    "fontWeight":"600"
                }
            ),
            html.Div(id="last-update-label",
                     style={"fontSize":"11px","color":"#999","alignSelf":"center"}),
        ], style={"display":"flex","gap":"12px","alignItems":"center","flexWrap":"wrap"}),
    ], style={
        "display":"flex","justifyContent":"space-between","alignItems":"center",
        "flexWrap":"wrap","gap":"10px",
        "padding":"12px 20px","background":CARD_BG,"borderBottom":"2px solid #e9eaef"
    }),

    # ── Вкладки + группировка тепловой карты ──────
    html.Div([
        html.Div(id="tab-nav-container", children=tab_buttons("overview")),
        html.Div(
            dcc.Dropdown(
                id="dropdown-groupby",
                options=[
                    {"label": "Группировка: месторождение", "value": "field"},
                    {"label": "Группировка: режим работы",  "value": "mode"},
                ],
                value="field", clearable=False,
                style={"width": "260px", "fontSize": "13px"},
            ),
            id="groupby-container", style={"display": "none"},
        ),
    ], style={
        "display": "flex", "justifyContent": "space-between", "alignItems": "center",
        "flexWrap": "wrap", "gap": "10px", "padding": "12px 20px 0",
    }),

    # ── Тело: боковая панель скважин + контент вкладки ─
    html.Div([

        # Левая колонка — список скважин
        html.Div([
            html.Div("Месторождения и скважины", style={
                "fontWeight": "700", "fontSize": "14px", "marginBottom": "8px", "padding": "0 2px",
                "color": TEXT_DARK,
            }),
            html.Div([
                svg_icon("search", color=TEXT_MUTED, size=15),
                dcc.Input(
                    id="well-search", type="text", placeholder="Поиск скважины…",
                    debounce=True,
                    style={
                        "border": "none", "outline": "none", "fontSize": "13px",
                        "marginLeft": "6px", "width": "100%", "background": "transparent",
                        "color": TEXT_DARK,
                    },
                ),
            ], style={
                "display": "flex", "alignItems": "center", "background": "#f5f6fa",
                "border": f"1px solid {BORDER_SOFT}", "borderRadius": "9px",
                "padding": "6px 10px", "marginBottom": "10px",
            }),
            html.Div(id="sidebar-container", style={
                "flex": 1, "minHeight": "0", "overflowY": "auto", "paddingRight": "4px",
            }),
        ], className="sidebar-col", style={
            "width": "270px", "minWidth": "150px", "background": CARD_BG,
            "border": "1px solid #e9eaef", "borderRadius": "14px", "padding": "12px",
            "boxShadow": "0 4px 16px rgba(20,23,28,0.06)",
            "display": "flex", "flexDirection": "column"
        }),

        # Правая колонка — контент активной вкладки
        html.Div(id="page-content", style={"flex": "1", "minWidth": "0", "display": "flex", "flexDirection": "column"}),

    ], className="body-row", style={"display": "flex", "gap": "14px", "padding": "12px 20px 20px", "alignItems": "stretch", "flex": "1", "minHeight": "0"}),

    # ── Модальное окно с деталями по скважине ─────
    html.Div(id="well-detail-panel", style={"display": "none"}),

], style={"background":BG,"fontFamily":"Arial, sans-serif","height":"100vh", "overflow": "auto", "display": "flex", "flexDirection": "column"})


# ════════════════════════════════════════════════
#  CALLBACKS
# ════════════════════════════════════════════════

@app.callback(
    Output("store-period", "data"),
    [Input(f"btn-period-{p}", "n_clicks") for p in PERIODS],
    State("store-period", "data"),
    # ← убрали prevent_initial_call=True
)
def update_period(*args):
    """Запоминает выбранный период (и форсирует пересчёт при первой загрузке)."""
    n_clicks_list = args[:-1]
    current = args[-1] or "1м"

    if not any(n_clicks_list):          # это самая первая загрузка, ни одна кнопка ещё не нажата
        return current                   # явно переустанавливаем "1м" — это и есть форс-триггер

    ctx = callback_context
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    return btn_id.replace("btn-period-", "")



@app.callback(
    Output("period-buttons-container", "children"),
    Input("store-period", "data"),
)
def refresh_period_buttons(period):
    return period_buttons(period)


@app.callback(
    Output("store-active-tab", "data"),
    [Input(f"btn-tab-{val}", "n_clicks") for val, _, _ in TABS],
    State("store-active-tab", "data"),
    # prevent_initial_call=True,   ← убрать
)
def switch_tab(*args):
    """Запоминает активную вкладку (и форсирует пересчёт при первой загрузке)."""
    n_clicks_list = args[:-1]
    current = args[-1] or "overview"

    if not any(n_clicks_list):          # первая загрузка — ни одна кнопка ещё не нажата
        return current                   # явно переустанавливаем восстановленное значение — форс-триггер

    ctx = callback_context
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    return btn_id.replace("btn-tab-", "")



@app.callback(
    [Output("tab-nav-container", "children"), Output("groupby-container", "style")],
    Input("store-active-tab", "data"),
)
def refresh_tab_nav(tab):
    groupby_style = {"display": "block", "width": "260px"} if tab == "heatmap" else {"display": "none"}
    return tab_buttons(tab), groupby_style


@app.callback(
    [
        Output("dropdown-field",      "options"),
        Output("page-content",        "children"),
        Output("last-update-label",   "children"),
    ],
    [
        Input("btn-refresh",      "n_clicks"),
        Input("dropdown-field",   "value"),
        Input("store-period",     "data"),
        Input("store-active-tab", "data"),
        Input("dropdown-groupby", "value"),
    ],
    State("store-consumption-tab", "data"),
)
def render_page(n_clicks, field, period, tab, group_by, consumption_tab):
    """Единая точка сборки контента активной вкладки."""
    df_all = load_data()

    fields = sorted(df_all["field"].unique())
    field_options = [{"label": "Все месторождения", "value": "ALL"}] + \
                    [{"label": f, "value": f} for f in fields]

    df = filter_df(df_all, field or "ALL", period or "1м")
    ts = datetime.now().strftime("Обновлено: %d.%m.%Y %H:%M:%S")

    if df.empty:
        empty_msg = html.Div("Нет данных за выбранный период", style={
            "padding": "60px 20px", "textAlign": "center", "color": TEXT_MUTED,
            "background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
        })
        return field_options, empty_msg, ts

    if tab == "heatmap":
        content = html.Div([
            html.Div("Тепловая карта скважин по УРЭ",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px", "color": TEXT_DARK}),
            html.Div("Размер блока — добыча жидкости, цвет — отклонение УРЭ факт от расчёта",
                     style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(id="graph-heatmap", figure=make_heatmap(df, group_by or "field"),
                       config={"displayModeBar": False, "responsive": False}),
        ], style={
            "background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
            "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"
        })
        return field_options, content, ts

    if tab == "trends":
        insights = generate_insights(df)
        insight_rows = []
        for it in insights:
            insight_rows.append(html.Div([
                html.Div(svg_icon(it["icon"], color=it["color"], size=17), style={
                    "width": "34px", "height": "34px", "borderRadius": "10px",
                    "background": with_alpha(it["color"], 0.14), "display": "flex",
                    "alignItems": "center", "justifyContent": "center", "flexShrink": "0",
                }),
                html.Div([
                    html.Div(it["title"], style={"fontSize": "13px", "fontWeight": "700", "color": TEXT_DARK}),
                    html.Div(it["text"], style={"fontSize": "12px", "color": TEXT_MUTED, "marginTop": "2px", "lineHeight": "1.4"}),
                ], style={"marginLeft": "10px", "flex": "1"}),
                svg_icon("chevron", color=TEXT_MUTED, size=14),
            ], style={
                "display": "flex", "alignItems": "flex-start", "padding": "12px 4px",
                "borderBottom": f"1px solid {BORDER_SOFT}",
            }))

        insights_card = html.Div([
            html.Div([svg_icon("bulb", color=GREEN_DARK, size=16),
                      html.Span("Предложения и рекомендации",
                                style={"fontWeight": "700", "fontSize": "14px", "marginLeft": "8px", "color": TEXT_DARK}),
                      html.Span("AI-аналитика", style={
                          "marginLeft": "auto", "fontSize": "10px", "fontWeight": "700", "color": GREEN_DARK,
                          "background": with_alpha(GREEN_DARK, 0.12), "padding": "3px 9px", "borderRadius": "9px",
                      })],
                     style={"display": "flex", "alignItems": "center", "marginBottom": "6px"}),
            html.Div(insight_rows, style={"marginTop": "4px"}),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "16px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)", "marginBottom": "12px"})

        charts_row = html.Div([
            html.Div([
                html.Div("Динамика УРЭ факт/расчёт", style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
                dcc.Graph(figure=make_trend(df), config={"displayModeBar": False}),
            ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                      "padding": "14px", "flex": "1", "minWidth": "320px",
                      "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),
            html.Div([
                html.Div("Потребление электроэнергии по дням", style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
                dcc.Graph(figure=make_consumption_trend(df), config={"displayModeBar": False}),
            ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                      "padding": "14px", "flex": "1", "minWidth": "320px",
                      "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),
        ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "marginBottom": "12px"})

        overconsumption_card = html.Div([
            html.Div("ТОП по перерасходу электроэнергии", style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
            html.Div("Скв. | сверх плана, тыс.кВт·ч", style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(figure=make_overconsumption_bar(df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"})

        return field_options, html.Div([insights_card, charts_row, overconsumption_card]), ts

    if tab == "rating":
        rating_card = html.Div([
            html.Div("Рейтинг скважин фонда", style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
            html.Div("Сортировка и фильтрация по любому столбцу",
                     style={"fontSize": "11px", "color": TEXT_MUTED, "marginBottom": "10px"}),
            build_rating_table(df),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)", "marginBottom": "12px"})

        # upload_card = html.Div([
        #     html.Div([svg_icon("upload", color=GREEN_DARK, size=16),
        #               html.Span("Мероприятия по снижению УРЭ",
        #                         style={"fontWeight": "700", "fontSize": "14px", "marginLeft": "8px", "color": TEXT_DARK})],
        #              style={"display": "flex", "alignItems": "center", "marginBottom": "4px"}),
        #     html.Div(
        #         "Загрузите сюда единый Excel-файл со всеми мероприятиями по фонду — колонки: "
        #         "«Скважина", "Мероприятие", "Рекомендации", "Возможная экономия электроэнергии, кВт». "
        #         "После загрузки они появятся на листе «Мероприятия» в карточке каждой скважины "
        #         f"(файл сохраняется как {MEROPRIYATIYA_PATH}, доступен по ссылке /meropriyatiya.xlsx).",
        #         style={"fontSize": "12px", "color": TEXT_MUTED, "marginBottom": "10px", "lineHeight": "1.5"},
        #     ),
        #     dcc.Upload(
        #         id="upload-meropriyatiya",
        #         children=html.Div([
        #             svg_icon("upload", color=GREY_LIGHT, size=22),
        #             html.Div("Перетащите файл сюда или нажмите, чтобы выбрать (.xlsx)",
        #                      style={"fontSize": "12px", "color": TEXT_MUTED, "marginTop": "6px"}),
        #         ], style={"textAlign": "center", "padding": "22px"}),
        #         style={
        #             "border": f"2px dashed {BORDER_SOFT}", "borderRadius": "12px",
        #             "background": "#f7f8fb", "cursor": "pointer",
        #         },
        #         multiple=False,
        #     ),
        #     html.Div(id="upload-meropriyatiya-status", style={"fontSize": "12px", "marginTop": "8px", "color": GREEN_DARK}),
        # ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
        #           "padding": "16px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"})

        return field_options, html.Div([rating_card]), ts

    # ── Вкладка «Обзор» ────────────────────────────
    e_fact  = int(df["electricity_fact"].sum())
    e_plan  = int(df["electricity_plan"].sum())
    liquid  = int(df["liquid"].sum())
    oil     = int(df["oil"].sum())
    ure_f   = df["ure_fact"].mean()
    ure_p   = df["ure_plan"].mean()
    dev     = round((ure_f - ure_p) / ure_p * 100, 2) if ure_p else 0

    def fmt(n): return f"{n:,}".replace(",", " ")

    upload_consumption_card = html.Div([
        html.Div([
                  html.Span("Загрузить данные по энергопотреблению",
                            style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK})],
                 style={"display": "flex", "alignItems": "center", "marginBottom": "4px"}),
        html.Div(
            "Внимание: дублирующиеся данные будут перезаписаны",
            style={"fontSize": "12px", "color": ORANGE, "marginBottom": "10px", "lineHeight": "1.5"},
        ),
        dcc.Upload(
            id="upload-consumption",
            children=html.Div([
                svg_icon("upload", color=GREY_LIGHT, size=22),
                html.Div("Перенесите файл сюда или нажмите, чтобы выбрать (.xlsx)",
                         style={"fontSize": "12px", "color": TEXT_MUTED, "marginLeft": "10px"}),  # ← marginLeft вместо marginTop
            ], style={
                "display": "flex", "flexDirection": "row",     # ← row вместо column
                "alignItems": "center", "justifyContent": "center",
                "height": "100%",
            }),
            style={
                "border": f"2px dashed {BORDER_SOFT}", "borderRadius": "12px",
                "background": "#f7f8fb", "cursor": "pointer",
                "height": "50px",   # при горизонтальном расположении можно сделать зону ниже
                "width": "100%",
            },
            multiple=False,
        ),

        html.Div(id="upload-consumption-status", style={"fontSize": "12px", "marginTop": "8px", "color": GREEN_DARK}),
    ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
              "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)", "marginBottom": "12px"})

    kpi_row = html.Div([
        kpi_card("bolt", "Электроэнергия, кВт·ч",      fmt(e_fact), color=PURPLE),
        kpi_card("plug", "Электроэнергия расч., кВт·ч", fmt(e_plan), color=BLUE),
        kpi_card("droplet", "Добыча жидкости, м³",          fmt(liquid), color=TEAL),
        kpi_card("barrel", "Добыча нефти, т",              fmt(oil), color=YELLOW),
        kpi_card("gauge", "УРЭ факт, кВт·ч/м³",           f"{ure_f:.2f}", color=GREEN_DARK),
        kpi_card("clipboard", "УРЭ расч., кВт·ч/м³",          f"{ure_p:.2f}", color=BLUE),
        kpi_card("trend", "Отклонение УРЭ, %",             f"{dev:.2f}", accent=True),
    ], style={"display": "flex", "gap": "10px", "flexWrap": "wrap", "marginBottom": "12px"})

    middle_row = html.Div([
        html.Div([
            html.Div("Анализ энергопотребления",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px", "color": TEXT_DARK}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": TEXT_MUTED, "marginBottom": "6px"}),
            dcc.Graph(figure=make_gauge(ure_f, df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "flex": "1", "minWidth": "250px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),

        html.Div([
            html.Div("ТОП лучших скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px", "color": TEXT_DARK}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": TEXT_MUTED, "marginBottom": "6px"}),
            dcc.Graph(figure=make_top_best(df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "flex": "1", "minWidth": "250px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),

        html.Div([
            html.Div("ТОП худших скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px", "color": TEXT_DARK}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": TEXT_MUTED, "marginBottom": "6px"}),
            dcc.Graph(figure=make_top_worst(df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "flex": "1", "minWidth": "250px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),

        html.Div([
            html.Div("Режимы работы скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "22px", "color": TEXT_DARK}),
            mode_panel(df),
        ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
                  "padding": "14px", "flex": "1", "minWidth": "250px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),

        html.Div(build_quick_summary(df), style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
               "padding": "14px", "flex": "1", "minWidth": "250px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)"}),
    ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "marginBottom": "12px"})

    # consumption_row = html.Div([
    #     html.Div([
    #         html.Span("Потребление электроэнергии, ",
    #                   style={"fontWeight": "700", "fontSize": "14px", "color": TEXT_DARK}),
    #         html.Span("тыс.кВт·ч", style={"fontWeight": "900", "fontSize": "14px", "color": TEXT_DARK}),
    #     ], style={"marginBottom": "12px"}),
    #     dcc.Graph(figure=make_consumption(df), config={"displayModeBar": False}),
    # ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
    #           "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)", "marginBottom": "12px"})

    consumption_row = html.Div([
        dcc.Tabs([
            dcc.Tab(label="Суммарное энергопотребление", value="period", style=TAB_STYLE, selected_style=TAB_SELECTED_STYLE,
                    children=html.Div(
                        dcc.Graph(figure=make_consumption_trend(df),
                                  config={"displayModeBar": False, "responsive": False}),
                        style={"paddingTop": "10px"},
                    )),
            dcc.Tab(label="По скважинам", value="wells", style=TAB_STYLE, selected_style=TAB_SELECTED_STYLE,
                    children=html.Div(
                        dcc.Graph(figure=make_consumption(df), config={"displayModeBar": False, "responsive": False}),
                        style={"paddingTop": "10px"},
                    )),
        ], id="consumption-tabs", value=consumption_tab or "period", persistence=True, persistence_type="local", style={"borderBottom": f"1px solid {BORDER_SOFT}"}),
    ], style={"background": CARD_BG, "border": "1px solid #e9eaef", "borderRadius": "14px",
              "padding": "14px", "boxShadow": "0 4px 16px rgba(20,23,28,0.06)", "marginBottom": "12px"})

    content = html.Div([kpi_row, middle_row, consumption_row, upload_consumption_card], style={"display": "flex", "flexDirection": "column", "flex": "1", "minHeight": "0"})
    return field_options, content, ts


# ── Боковая панель (месторождения → скважины + поиск) ─
@app.callback(
    Output("sidebar-container", "children"),
    [
        Input("btn-refresh",    "n_clicks"),
        Input("dropdown-field", "value"),
        Input("store-period",   "data"),
        Input("well-search",    "value"),
    ],
)
def render_sidebar(n_clicks, field, period, search_term):
    df_all = load_data()
    df = filter_df(df_all, field or "ALL", period or "1м")
    return build_sidebar_wells(df, search_term)


# ── Выбор скважины из бокового списка ─────────────
@app.callback(
    Output("store-selected-well", "data", allow_duplicate=True),
    Input({"type": "well-card", "index": ALL}, "n_clicks"),
    prevent_initial_call=True,
)
def select_well_sidebar(n_clicks_list):
    ctx = callback_context
    if not ctx.triggered or not any(n_clicks_list):
        raise PreventUpdate
    trig = ctx.triggered[0]["prop_id"].split(".")[0]
    well_id = json.loads(trig)["index"]
    return well_id

@app.callback(
    Output("store-consumption-tab", "data"),
    Input("consumption-tabs", "value"),
    prevent_initial_call=True,
)
def remember_consumption_tab(value):
    return value


# ── Выбор скважины кликом по тепловой карте ───────
@app.callback(
    Output("store-selected-well", "data", allow_duplicate=True),
    Input("graph-heatmap", "clickData"),
    prevent_initial_call=True,
)
def select_well_heatmap(click_data):
    if not click_data:
        raise PreventUpdate
    label = click_data["points"][0].get("label")
    df_all = load_data()
    if label is not None and label in df_all["well"].astype(str).unique():
        return label
    raise PreventUpdate


# ── Закрытие окна деталей ─────────────────────────
@app.callback(
    Output("store-selected-well", "data", allow_duplicate=True),
    Input("btn-close-detail", "n_clicks"),
    prevent_initial_call=True,
)
def close_well_detail(n_clicks):
    return None


# ── Сохранение загруженного Excel с мероприятиями ─
@app.callback(
    Output("upload-meropriyatiya-status", "children"),
    Input("upload-meropriyatiya", "contents"),
    State("upload-meropriyatiya", "filename"),
    prevent_initial_call=True,
)
def save_uploaded_meropriyatiya(contents, filename):
    if contents is None:
        raise PreventUpdate
    try:
        _content_type, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)
        os.makedirs(MEROPRIYATIYA_DIR, exist_ok=True)
        with open(MEROPRIYATIYA_PATH, "wb") as f:
            f.write(decoded)
        check_df = pd.read_excel(MEROPRIYATIYA_PATH)
        return f"✅ «{filename}» загружен, строк: {len(check_df)}. Откройте карточку скважины → лист «Мероприятия»."
    except Exception as e:
        return f"⚠️ Не удалось прочитать файл: {e}"

@app.callback(
    [Output("upload-consumption-status", "children"),
     Output("store-period", "data", allow_duplicate=True)],   # триггерит перерисовку без нажатия "Обновить"
    Input("upload-consumption", "contents"),
    [State("upload-consumption", "filename"), State("store-period", "data")],
    prevent_initial_call=True,
)
def save_uploaded_consumption(contents, filename, current_period):
    if contents is None:
        raise PreventUpdate
    try:
        _content_type, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)

        new_df = pd.read_excel(io.BytesIO(decoded), sheet_name=EXCEL_SHEET)
        new_df.columns = new_df.columns.str.strip().str.lower()
        required = {"timestamp", "field", "well", "electricity_fact", "electricity_plan",
                    "liquid", "oil", "ure_fact", "ure_plan"}
        missing = required - set(new_df.columns)
        if missing:
            return f"⚠️ Отсутствуют колонки: {missing}", dash.no_update

        new_df["timestamp"] = pd.to_datetime(new_df["timestamp"])
        new_df["well"] = new_df["well"].astype(str)

        os.makedirs(ADDITIONAL_DATA_DIR, exist_ok=True)
        if os.path.exists(ADDITIONAL_DATA_PATH):
            existing = pd.read_excel(ADDITIONAL_DATA_PATH, sheet_name=EXCEL_SHEET)
            existing.columns = existing.columns.str.strip().str.lower()
            combined = pd.concat([existing, new_df], ignore_index=True)
        else:
            combined = new_df

        combined = combined.drop_duplicates(subset=["timestamp", "well"], keep="last")
        combined.to_excel(ADDITIONAL_DATA_PATH, sheet_name=EXCEL_SHEET, index=False)

        return (f"✅ «{filename}» добавлен: {len(new_df)} строк. Всего накоплено: {len(combined)} строк.",
                current_period)  # тот же период — форсирует перерисовку через существующий Input
    except Exception as e:
        return f"⚠️ Не удалось прочитать файл: {e}", dash.no_update


# @app.callback(
#     Output("rating-table", "data"),
#     Input("rating-table", "sort_by"),
#     State("rating-table", "data"),
#     prevent_initial_call=True,
# )
# def sort_rating_table(sort_by, current_data):
#     if not sort_by:
#         raise PreventUpdate
#     d = pd.DataFrame(current_data)
#     col = sort_by[0]["column_id"]
#     ascending = sort_by[0]["direction"] == "asc"
#     d = d.sort_values(col, ascending=ascending, kind="mergesort")
#     return d.to_dict("records")

@app.callback(
    Output('rating-table', 'data'),
    Input('rating-table', 'sort_by'),
    State('rating-table', 'data')
)
def sort_table(sort_by, data):
    if data is None:
        return []

    # Преобразуем в DataFrame
    df = pd.DataFrame(data)

    if sort_by:
        # Сортируем
        sort_col = sort_by[0]['column_id']
        sort_asc = sort_by[0]['direction'] == 'asc'
        df = df.sort_values(by=sort_col, ascending=sort_asc)

    return df.to_dict('records')

app.clientside_callback(
    """
    function(tableData) {
        setTimeout(function() {
            document.querySelectorAll('#rating-table .dash-filter input[type="text"]').forEach(function(el) {
                el.placeholder = 'фильтр…';
            });
        }, 200);
        return window.dash_clientside.no_update;
    }
    """,
    Output("dummy-filter-fix", "data"),
    Input("rating-table", "data"),
)


# ── Рендер модального окна с деталями скважины ────
@app.callback(
    [Output("well-detail-panel", "children"), Output("well-detail-panel", "style")],
    Input("store-selected-well", "data"),
    State("store-period", "data"),
)
def render_well_panel(well, period):
    hidden_style = {"display": "none"}
    if not well:
        return None, hidden_style

    df_all = load_data()
    delta = PERIODS.get(period or "1м", timedelta(days=30))
    cutoff = df_all["timestamp"].max() - delta
    df_period = df_all[df_all["timestamp"] >= cutoff]

    if well not in df_period["well"].astype(str).unique():
        return None, hidden_style

    card = html.Div(
        build_well_detail(df_period, well),
        style={
            "background": CARD_BG, "borderRadius": "10px", "padding": "22px",
            "width": "min(92vw, 900px)", "maxHeight": "88vh", "overflowY": "auto",
            "boxShadow": "0 8px 30px rgba(0,0,0,0.25)",
        },
    )
    visible_style = {
        "display": "flex", "position": "fixed", "top": 0, "left": 0, "right": 0, "bottom": 0,
        "background": "rgba(0,0,0,0.45)", "alignItems": "center", "justifyContent": "center",
        "zIndex": 1000,
    }
    return card, visible_style


# ════════════════════════════════════════════════
if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=8050)
