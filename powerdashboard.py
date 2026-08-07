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
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

import dash
from dash import dcc, html, Input, Output, State, callback_context, ALL
from dash.exceptions import PreventUpdate
import plotly.graph_objects as go

from dash import Dash

# ════════════════════════════════════════════════
#  ПУТЬ К ФАЙЛУ — поменяйте на свой
# ════════════════════════════════════════════════
EXCEL_PATH = "data.xlsx"        # абсолютный или относительный путь
EXCEL_SHEET = "data"            # имя листа

# ── Цвета ────────────────────────────────────────
GREEN_DARK   = "rgba(0, 150, 57, 1)"
GREEN_MID    = "rgba(0, 150, 57, 0.7)"
GREEN_LIGHT  = "rgba(0, 150, 57, 0.5)"
GREEN_TRANSP = "rgba(0, 150, 57, 0.3)"
GREEN        = "rgba(155, 189, 30, 1)"
YELLOW       = "rgba(234, 170, 0, 0.8)"
YELLOW_LIGHT = "rgba(234, 170, 0, 0.5)"
YELLOW_TRANSP = "rgba(234, 170, 0, 0.3)"
ORANGE        = "rgba(237, 109, 27, 1)"
ORANGE_MID    = "rgba(237, 109, 27, 0.7)"
ORANGE_LIGHT  = "rgba(237, 109, 27, 0.5)"
ORANGE_TRANSP = "rgba(237, 109, 27, 0.3)"
GREY_DARK    = "#546e7a"
GREY_LIGHT   = "#b0bec5"
BG           = "#f4f6f8"
CARD_BG      = "#ffffff"

MODES = ["в работе", "в накоплении", "в простое", "в бездействии"]
MODE_COLORS = {
    "в работе":      GREEN_DARK,
    "в накоплении":  GREEN,
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
    rows = []
    # записи каждые 8 часов за 30 дней
    for days_back in range(30, -1, -1):
        for hour in [0, 8, 16]:
            ts = now - timedelta(days=days_back, hours=now.hour) + timedelta(hours=hour)
            for field, wlist in wells_by_field.items():
                for w in wlist:
                    base = np.random.uniform(40, 95)
                    ure  = np.random.uniform(6, 44)
                    mode = np.random.choice(MODES, p=[0.55, 0.02, 0.03, 0.40])
                    rows.append({
                        "timestamp":        ts,
                        "field":            field,
                        "well":             w,
                        "electricity_fact": round(base * 1000),
                        "electricity_plan": round(base * 0.97 * 1000),
                        "liquid":           round(np.random.uniform(2000, 5000)),
                        "oil":              round(np.random.uniform(200, 600)),
                        "ure_fact":         round(ure, 2),
                        "ure_plan":         round(ure * 0.96, 2),
                        "mode":             mode,
                    })
    return pd.DataFrame(rows)


def load_data() -> pd.DataFrame:
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
        font=dict(size=13, color=GREY_DARK),
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
        height=260,
        margin=dict(l=20, r=20, t=5, b=30),
        paper_bgcolor=CARD_BG,
        plot_bgcolor="rgba(0,0,0,0)",
        font={"family": "Arial"},
        annotations=annotations,
        xaxis=dict(visible=False, range=[0, 1]),
        yaxis=dict(visible=False, range=[0, 1]),
    )
    return fig


def make_top_best(df: pd.DataFrame):
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
        height=260, margin=dict(l=50,r=60,t=10,b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(range=[0, agg.max()*1.3], showgrid=False, zeroline=False, visible=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_top_worst(df: pd.DataFrame):
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
        height=260, margin=dict(l=50,r=60,t=10,b=30),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        xaxis=dict(range=[0, agg.max()*1.3], showgrid=False, zeroline=False, visible=False),
        yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_consumption(df: pd.DataFrame):
    agg = df.groupby("well").agg(
        fact=("electricity_fact","sum"),
        plan=("electricity_plan","sum")
    ).reset_index().sort_values("well")
    agg["fact_k"] = (agg["fact"] / 1000).round(0).astype(int)
    agg["plan_k"] = (agg["plan"] / 1000).round(0).astype(int)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="Фактическое потребление", x=agg["well"], y=agg["fact_k"],
        marker_color=GREY_DARK,
        text=agg["fact_k"], textposition="outside", textfont=dict(size=9),
        hovertemplate="%{x}: %{y} тыс.кВт·ч<extra></extra>"
    ))
    fig.add_trace(go.Bar(
        name="Расчётное потребление", x=agg["well"], y=agg["plan_k"],
        marker_color=GREY_LIGHT,
        text=agg["plan_k"], textposition="outside", textfont=dict(size=9),
        hovertemplate="%{x}: %{y} тыс.кВт·ч<extra></extra>"
    ))
    ymax = max(agg["fact_k"].max(), agg["plan_k"].max()) * 1.25 if len(agg) else 100
    fig.update_layout(
        barmode="group", bargap=0.15, bargroupgap=0.02,
        height=280, margin=dict(l=40,r=150,t=30,b=40),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(range=[0, ymax], gridcolor="#eeeeee", title="тыс.кВт·ч", title_font=dict(size=11)),
        xaxis=dict(tickfont=dict(size=10)),
        legend=dict(orientation="v", x=1.01, y=1, font=dict(size=11)),
        font={"family":"Arial"}
    )
    return fig


def make_trend(df: pd.DataFrame):
    """Линейный тренд УРЭ факт по времени."""
    agg = df.groupby("timestamp")["ure_fact"].mean().reset_index()
    fig = go.Figure(go.Scatter(
        x=agg["timestamp"], y=agg["ure_fact"],
        mode="lines+markers", line=dict(color=ORANGE, width=2),
        marker=dict(size=4), name="УРЭ факт",
        hovertemplate="%{x|%d.%m %H:%M}<br>УРЭ: %{y:.2f} кВт·ч/м³<extra></extra>"
    ))
    fig.update_layout(
        height=180, margin=dict(l=50,r=20,t=20,b=40),
        paper_bgcolor=CARD_BG, plot_bgcolor=CARD_BG,
        yaxis=dict(gridcolor="#eeeeee", title="кВт·ч/м³", title_font=dict(size=10)),
        xaxis=dict(gridcolor="#eeeeee"),
        font={"family":"Arial"}
    )
    return fig


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
#  ЛЕВАЯ ПАНЕЛЬ — СПИСОК СКВАЖИН (кликабельный)
# ════════════════════════════════════════════════

def build_sidebar_wells(df: pd.DataFrame):
    if df.empty:
        return html.Div("Нет данных", style={"fontSize": "12px", "color": "#999", "padding": "8px"})

    latest = df.sort_values("timestamp").groupby("well").last().reset_index()[["well", "field"]]
    agg = df.groupby("well").agg(
        ure_fact=("ure_fact", "mean"),
        ure_plan=("ure_plan", "mean"),
        liquid=("liquid", "sum"),
    ).reset_index()
    agg["dev"] = np.where(agg["ure_plan"] != 0,
                           (agg["ure_fact"] - agg["ure_plan"]) / agg["ure_plan"] * 100, 0)
    agg = agg.merge(latest, on="well", how="left").sort_values("dev", ascending=False)

    cards = []
    for _, row in agg.iterrows():
        color, _label = well_status(row["dev"])
        cards.append(html.Div([
            html.Div([
                html.Span(style={
                    "display": "inline-block", "width": "8px", "height": "8px",
                    "borderRadius": "50%", "background": color, "marginRight": "7px",
                    "flexShrink": "0",
                }),
                html.Span(f"Скв. {row['well']}", style={"fontWeight": "700", "fontSize": "13px"}),
                html.Span(f"{row['dev']:+.1f}%", style={
                    "marginLeft": "auto", "fontSize": "11px", "fontWeight": "700", "color": color
                }),
            ], style={"display": "flex", "alignItems": "center"}),
            html.Div(f"{row['field']} · УРЭ {row['ure_fact']:.1f} кВт·ч/м³",
                     style={"fontSize": "11px", "color": "#888", "marginTop": "2px",
                            "whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis"}),
        ],
            id={"type": "well-card", "index": str(row["well"])}, n_clicks=0,
            style={
                "background": "#fff", "border": "1px solid #e0e0e0", "borderRadius": "7px",
                "padding": "8px 10px", "marginBottom": "6px", "cursor": "pointer",
            }
        ))
    return html.Div(cards)


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


def build_well_detail(df_period: pd.DataFrame, well: str):
    d = df_period[df_period["well"].astype(str) == str(well)]
    if d.empty:
        return html.Div("Нет данных по скважине за выбранный период.")

    last = d.sort_values("timestamp").iloc[-1]
    ure_f, ure_p = d["ure_fact"].mean(), d["ure_plan"].mean()
    dev = (ure_f - ure_p) / ure_p * 100 if ure_p else 0
    color, label = well_status(dev)
    mode = last["mode"] if "mode" in last and pd.notna(last["mode"]) else "—"

    def fmt(n):
        return f"{n:,.0f}".replace(",", " ")

    header = html.Div([
        html.Div([
            html.Span(f"Скважина {well}", style={"fontSize": "18px", "fontWeight": "700"}),
            html.Span(last["field"], style={"fontSize": "12px", "color": "#888", "marginLeft": "10px"}),
            html.Span(label, style={
                "marginLeft": "10px", "fontSize": "11px", "fontWeight": "700", "color": "white",
                "background": color, "padding": "2px 10px", "borderRadius": "10px",
            }),
            html.Span(f"Режим: {mode}", style={"fontSize": "11px", "color": "#888", "marginLeft": "10px"}),
        ]),
        html.Button("✕", id="btn-close-detail", n_clicks=0, style={
            "border": "none", "background": "none", "fontSize": "20px",
            "cursor": "pointer", "color": "#888", "lineHeight": "1",
        }),
    ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "marginBottom": "14px"})

    kpis = html.Div([
        kpi_card("⚡", "ЭЭ факт, кВт·ч", fmt(d["electricity_fact"].sum())),
        kpi_card("🔌", "ЭЭ расч, кВт·ч", fmt(d["electricity_plan"].sum())),
        kpi_card("💧", "Жидкость, м³", fmt(d["liquid"].sum())),
        kpi_card("🛢️", "Нефть, т", fmt(d["oil"].sum())),
        kpi_card("⚡", "УРЭ факт, кВт·ч/м³", f"{ure_f:.2f}"),
        kpi_card("📈", "Откл. УРЭ, %", f"{dev:+.1f}", accent=True),
    ], style={"display": "flex", "gap": "8px", "flexWrap": "wrap", "marginBottom": "14px"})

    charts = html.Div([
        html.Div([
            html.Div("Динамика УРЭ и дебита жидкости",
                     style={"fontWeight": "700", "fontSize": "13px", "marginBottom": "4px"}),
            dcc.Graph(figure=build_well_trend_fig(df_period, well), config={"displayModeBar": False}),
        ], style={"flex": "1.4", "minWidth": "300px"}),
        html.Div([
            html.Div("Скважина vs среднее по месторождению",
                     style={"fontWeight": "700", "fontSize": "13px", "marginBottom": "4px"}),
            dcc.Graph(figure=build_well_compare_bar(df_period, well), config={"displayModeBar": False}),
        ], style={"flex": "1", "minWidth": "220px"}),
    ], style={"display": "flex", "gap": "16px", "flexWrap": "wrap"})

    return html.Div([header, kpis, charts])


# ════════════════════════════════════════════════
#  ТЕПЛОВАЯ КАРТА (Treemap: группа → скважина)
# ════════════════════════════════════════════════

HEAT_COLORSCALE = [
    [0.0,  GREEN_DARK],
    [0.30, GREEN],
    [0.55, YELLOW],
    [0.78, ORANGE_MID],
    [1.0,  ORANGE],
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
        f"<b>{w}</b><br>ЭЭ {e:,.0f}<br>УРЭ {uf:.1f}/{up:.1f}".replace(",", " ")
        for w, e, uf, up in zip(agg["well"].astype(str), agg["electricity"], agg["ure_fact"], agg["ure_plan"])
    ]
    hover = [f"{g}<extra></extra>" for g in groups[group_col].astype(str)] + [
        f"Скв. {w}<br>УРЭ факт: {uf:.2f} кВт·ч/м³<br>УРЭ расч: {up:.2f} кВт·ч/м³<br>Откл: {dv:+.1f}%<extra></extra>"
        for w, uf, up, dv in zip(agg["well"].astype(str), agg["ure_fact"], agg["ure_plan"], agg["dev"])
    ]

    fig = go.Figure(go.Treemap(
        labels=labels, parents=parents, values=values, branchvalues="total",
        text=text, texttemplate="%{text}", textfont=dict(size=11, color="white"),
        hovertemplate=hover,
        marker=dict(
            colors=colors, colorscale=HEAT_COLORSCALE, cmin=-10, cmax=35,
            line=dict(width=1.5, color="white"),
            colorbar=dict(title="Откл УРЭ,%", thickness=12, len=0.8),
        ),
        pathbar=dict(visible=True, textfont=dict(size=11)),
    ))
    fig.update_layout(
        height=560, margin=dict(l=5, r=5, t=30, b=5),
        paper_bgcolor=CARD_BG, font={"family": "Arial"},
    )
    return fig


# ════════════════════════════════════════════════
#  ВСПОМОГАТЕЛЬНЫЕ КОМПОНЕНТЫ
# ════════════════════════════════════════════════

def kpi_card(icon, title, value, accent=False):
    border    = f"2px solid {ORANGE}" if accent else "1px solid #e0e0e0"
    val_color = ORANGE if accent else "#1a1a2e"
    return html.Div([
        html.Span(icon, style={"fontSize":"20px","marginRight":"8px"}),
        html.Div([
            html.Div(title, style={"fontSize":"10px","color":"#777","lineHeight":"1.3"}),
            html.Div(value, style={"fontSize":"20px","fontWeight":"700","color":val_color}),
        ])
    ], style={
        "display":"flex","alignItems":"center",
        "background":CARD_BG,"border":border,"borderRadius":"8px",
        "padding":"10px 14px","flex":"1","minWidth":"130px",
        "boxShadow":"0 1px 4px rgba(0,0,0,0.07)"
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
        ], style={"background":"#f9f9f9","borderRadius":"8px","padding":"10px","textAlign":"center"})
        for m in MODES
    ], style={"display":"grid","gridTemplateColumns":"1fr 1fr","gap":"8px"})

    return html.Div([bar, grid])


def period_buttons(active="1м"):
    buttons = []
    for p in PERIODS:
        active_style = {
            "background": ORANGE, "color": "white",
            "border": f"1px solid {ORANGE}", "fontWeight": "700"
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


TABS = [("overview", "Обзор"), ("heatmap", "Тепловая карта")]


def tab_buttons(active="overview"):
    buttons = []
    for val, label in TABS:
        active_style = {
            "background": GREEN_DARK, "color": "white",
            "border": f"1px solid {GREEN_DARK}", "fontWeight": "700",
        }
        idle_style = {
            "background": CARD_BG, "color": GREY_DARK,
            "border": "1px solid #ddd", "fontWeight": "400",
        }
        style = {
            **(active_style if val == active else idle_style),
            "padding": "6px 18px", "borderRadius": "5px",
            "cursor": "pointer", "fontSize": "13px",
        }
        buttons.append(html.Button(label, id=f"btn-tab-{val}", n_clicks=0, style=style))
    return html.Div(buttons, style={"display": "flex", "gap": "8px"})


# ════════════════════════════════════════════════
#  APP LAYOUT
# ════════════════════════════════════════════════

# app = dash.Dash(__name__, title="Анализ энергопотребления скважин", suppress_callback_exceptions=True)
app = dash.Dash(__name__, title="Анализ энергопотребления скважин", suppress_callback_exceptions=True)
server = app.server

app.layout = html.Div([

    # ── Хранилище состояния ──────────────────────
    dcc.Store(id="store-period", data="1м"),
    dcc.Store(id="store-active-tab", data="overview"),
    dcc.Store(id="store-selected-well", data=None),

    # ── Шапка ────────────────────────────────────
    html.Div([
        html.Img(
        src="/assets/logo.png",
        style={
            "maxWidth": "100%",
            "height": "50px",
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
                style={"width":"220px","fontSize":"13px"}
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
        "padding":"12px 20px","background":CARD_BG,"borderBottom":"2px solid #e0e0e0"
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
            html.Div("Скважины", style={
                "fontWeight": "700", "fontSize": "14px", "marginBottom": "8px", "padding": "0 2px",
            }),
            html.Div(id="sidebar-container", style={
                "maxHeight": "76vh", "overflowY": "auto", "paddingRight": "4px",
            }),
        ], style={
            "width": "260px", "minWidth": "220px", "background": CARD_BG,
            "border": "1px solid #e0e0e0", "borderRadius": "8px", "padding": "12px",
            "boxShadow": "0 1px 4px rgba(0,0,0,0.07)", "alignSelf": "flex-start",
        }),

        # Правая колонка — контент активной вкладки
        html.Div(id="page-content", style={"flex": "1", "minWidth": "0"}),

    ], style={"display": "flex", "gap": "14px", "padding": "12px 20px 20px", "alignItems": "flex-start"}),

    # ── Модальное окно с деталями по скважине ─────
    html.Div(id="well-detail-panel", style={"display": "none"}),

], style={"background":BG,"fontFamily":"Arial, sans-serif","minHeight":"100vh"})


# ════════════════════════════════════════════════
#  CALLBACKS
# ════════════════════════════════════════════════

@app.callback(
    Output("store-period", "data"),
    [Input(f"btn-period-{p}", "n_clicks") for p in PERIODS],
    State("store-period", "data"),
    prevent_initial_call=True,
)
def update_period(*args):
    """Запоминает выбранный период."""
    ctx = callback_context
    if not ctx.triggered:
        return args[-1]
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    period = btn_id.replace("btn-period-", "")
    return period


@app.callback(
    Output("period-buttons-container", "children"),
    Input("store-period", "data"),
)
def refresh_period_buttons(period):
    return period_buttons(period)


@app.callback(
    Output("store-active-tab", "data"),
    [Input(f"btn-tab-{val}", "n_clicks") for val, _ in TABS],
    State("store-active-tab", "data"),
    prevent_initial_call=True,
)
def switch_tab(*args):
    """Запоминает активную вкладку."""
    ctx = callback_context
    if not ctx.triggered:
        return args[-1]
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
        Output("sidebar-container",   "children"),
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
)
def render_page(n_clicks, field, period, tab, group_by):
    """Единая точка сборки страницы: боковой список скважин + контент активной вкладки."""
    df_all = load_data()

    fields = sorted(df_all["field"].unique())
    field_options = [{"label": "Все месторождения", "value": "ALL"}] + \
                    [{"label": f, "value": f} for f in fields]

    df = filter_df(df_all, field or "ALL", period or "1м")
    ts = datetime.now().strftime("Обновлено: %d.%m.%Y %H:%M:%S")
    sidebar = build_sidebar_wells(df)

    if df.empty:
        empty_msg = html.Div("Нет данных за выбранный период", style={
            "padding": "60px 20px", "textAlign": "center", "color": "#999",
            "background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
        })
        return field_options, sidebar, empty_msg, ts

    if tab == "heatmap":
        content = html.Div([
            html.Div("Тепловая карта скважин по УРЭ",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px"}),
            html.Div("Размер блока — добыча жидкости, цвет — отклонение УРЭ факт от расчёта",
                     style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(id="graph-heatmap", figure=make_heatmap(df, group_by or "field"),
                       config={"displayModeBar": False}),
        ], style={
            "background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
            "padding": "14px", "boxShadow": "0 1px 4px rgba(0,0,0,0.07)",
        })
        return field_options, sidebar, content, ts

    # ── Вкладка «Обзор» ────────────────────────────
    e_fact  = int(df["electricity_fact"].sum())
    e_plan  = int(df["electricity_plan"].sum())
    liquid  = int(df["liquid"].sum())
    oil     = int(df["oil"].sum())
    ure_f   = df["ure_fact"].mean()
    ure_p   = df["ure_plan"].mean()
    dev     = round((ure_f - ure_p) / ure_p * 100, 2) if ure_p else 0

    def fmt(n): return f"{n:,}".replace(",", " ")

    kpi_row = html.Div([
        kpi_card("⚡", "Электроэнергия, кВт·ч",      fmt(e_fact)),
        kpi_card("🔌", "Электроэнергия расч., кВт·ч", fmt(e_plan)),
        kpi_card("💧", "Добыча жидкости, м³",          fmt(liquid)),
        kpi_card("🛢️", "Добыча нефти, т",              fmt(oil)),
        kpi_card("⚡", "УРЭ факт, кВт·ч/м³",           f"{ure_f:.2f}"),
        kpi_card("📋", "УРЭ расч., кВт·ч/м³",          f"{ure_p:.2f}"),
        kpi_card("📈", "Отклонение УРЭ, %",             f"{dev:.2f}", accent=True),
    ], style={"display": "flex", "gap": "10px", "flexWrap": "wrap", "marginBottom": "12px"})

    middle_row = html.Div([
        html.Div([
            html.Div("Анализ энергопотребления",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px"}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(figure=make_gauge(ure_f, df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
                  "padding": "14px", "flex": "1.1", "minWidth": "260px",
                  "boxShadow": "0 1px 4px rgba(0,0,0,0.07)"}),

        html.Div([
            html.Div("ТОП лучших скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px"}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(figure=make_top_best(df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
                  "padding": "14px", "flex": "1", "minWidth": "220px",
                  "boxShadow": "0 1px 4px rgba(0,0,0,0.07)"}),

        html.Div([
            html.Div("ТОП худших скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "2px"}),
            html.Div("Скв. | УРЭ факт, кВт·ч/м³",
                     style={"fontSize": "11px", "color": ORANGE, "marginBottom": "6px"}),
            dcc.Graph(figure=make_top_worst(df), config={"displayModeBar": False}),
        ], style={"background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
                  "padding": "14px", "flex": "1", "minWidth": "220px",
                  "boxShadow": "0 1px 4px rgba(0,0,0,0.07)"}),

        html.Div([
            html.Div("Режимы работы скважин",
                     style={"fontWeight": "700", "fontSize": "14px", "marginBottom": "12px"}),
            mode_panel(df),
        ], style={"background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
                  "padding": "14px", "flex": "0.9", "minWidth": "200px",
                  "boxShadow": "0 1px 4px rgba(0,0,0,0.07)"}),
    ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap", "marginBottom": "12px"})

    consumption_row = html.Div([
        html.Span("Потребление электроэнергии, ", style={"fontWeight": "700", "fontSize": "14px"}),
        html.Span("тыс.кВт·ч", style={"fontWeight": "900", "fontSize": "14px"}),
        dcc.Graph(figure=make_consumption(df), config={"displayModeBar": False}),
    ], style={"background": CARD_BG, "border": "1px solid #e0e0e0", "borderRadius": "8px",
              "padding": "14px", "boxShadow": "0 1px 4px rgba(0,0,0,0.07)"})

    content = html.Div([kpi_row, middle_row, consumption_row])
    return field_options, sidebar, content, ts


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
    app.run(debug=False, host='0.0.0.0', port=8050)
