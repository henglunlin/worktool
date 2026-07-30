
import os
import re
import tempfile
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import skrf as rf
import streamlit as st

# =========================================================
# Page Config
# =========================================================
st.set_page_config(page_title="SNP Viewer", layout="wide")
TODAY_MMDD = datetime.now().strftime("%m%d")

# =========================================================
# Session State
# =========================================================
def init_state(key, value):
    if key not in st.session_state:
        st.session_state[key] = value

init_state("applied_markers", [])
init_state("marker_text_buffer", "")
init_state("smith_filter_start", 0.0)
init_state("smith_filter_stop", 0.0)
init_state("smith_display_width_pct", 70)
init_state("smith_axis_limit", 1.25)
init_state("enable_png_export", False)

# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.1rem; padding-bottom: 2rem;}
    div[data-testid="stPopover"] > button {width: 100%;}
    .small-note {font-size: 0.88rem; color: #64748b;}
    .tool-title {font-weight: 700; color: #334155;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("SNP File Viewer")
st.caption("支援多 SNP 比對。Trace 選項已改成每個 SNP 檔案獨立選擇，避免所有檔案共用同一組 S-Parameter。")

# =========================================================
# Utility Functions
# =========================================================
def open_panel(label, key=None):
    """Use st.popover if available; fallback to st.expander for older Streamlit."""
    if hasattr(st, "popover"):
        return st.popover(label, use_container_width=True, key=key)
    return st.expander(label, expanded=False)


@st.cache_resource(show_spinner=False)
def read_touchstone_cached(file_name, file_bytes):
    """Cache Touchstone parsing so UI changes do not re-read the same SNP file."""
    suffix = os.path.splitext(file_name)[1] or ".s2p"
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    temp_file.write(file_bytes)
    temp_file.close()
    try:
        return rf.Network(temp_file.name)
    finally:
        try:
            os.remove(temp_file.name)
        except OSError:
            pass


def read_touchstone(uploaded_file):
    return read_touchstone_cached(uploaded_file.name, uploaded_file.getvalue())


def get_sparameter_list(nports):
    return [f"S{row}{col}" for row in range(1, nports + 1) for col in range(1, nports + 1)]


def parse_sparameter(sparam):
    match = re.fullmatch(r"S(\d+)(\d+)", str(sparam).upper())
    if not match:
        raise ValueError(f"Invalid S-Parameter: {sparam}")
    return int(match.group(1)) - 1, int(match.group(2)) - 1


def sort_sparam_key(sparam):
    row, col = parse_sparameter(sparam)
    return row, col


def get_reflection_sparams(sparams):
    return [s for s in sparams if parse_sparameter(s)[0] == parse_sparameter(s)[1]]


def safe_log10(value):
    value = np.maximum(np.abs(value), 1e-20)
    return np.log10(value)


def gamma_from_normalized_impedance(z_norm):
    return (z_norm - 1) / (z_norm + 1)


def normalized_impedance_from_gamma(gamma):
    gamma = np.asarray(gamma, dtype=complex)
    denominator = 1 - gamma
    denominator = np.where(np.abs(denominator) < 1e-20, 1e-20 + 0j, denominator)
    return (1 + gamma) / denominator


def normalized_admittance_from_z(z_norm):
    z_norm = np.asarray(z_norm, dtype=complex)
    denominator = np.where(np.abs(z_norm) < 1e-20, 1e-20 + 0j, z_norm)
    return 1 / denominator


def sanitize_sheet_name(name):
    base_name = os.path.splitext(os.path.basename(name))[0]
    base_name = re.sub(r"[\\/\*?:\[\]]", "_", base_name)
    return (base_name or "Sheet")[:31]


def make_unique_sheet_name(name, used_names):
    clean_name = sanitize_sheet_name(name)
    candidate = clean_name
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = f"{clean_name[:31-len(suffix)]}{suffix}"
        counter += 1
    used_names.add(candidate)
    return candidate


def parse_marker_input(marker_text):
    if not marker_text.strip():
        return []
    parts = re.split(r"[,;\s]+", marker_text.strip())
    markers = []
    for part in parts:
        if not part:
            continue
        try:
            markers.append(float(part))
        except ValueError:
            st.warning(f"Marker frequency parse fail, skipped: {part}")
    return sorted(set(markers))


def update_markers_from_input():
    st.session_state.applied_markers = parse_marker_input(st.session_state.marker_text_buffer)


def calculate_trace_data(network, sparam, data_type):
    row, col = parse_sparameter(sparam)
    if row >= network.nports or col >= network.nports:
        raise IndexError(f"{sparam} is not available for {network.nports}-port file")

    s_complex = network.s[:, row, col]
    mag = np.abs(s_complex)

    if data_type == "Magnitude dB":
        return 20 * safe_log10(mag), "Magnitude (dB)"
    if data_type == "Phase deg":
        return np.angle(s_complex, deg=True), "Phase (deg)"
    if data_type == "VSWR":
        if row != col:
            return np.full_like(mag, np.nan, dtype=float), "VSWR"
        gamma = np.clip(mag, 0, 0.999999)
        return (1 + gamma) / (1 - gamma), "VSWR"
    if data_type == "Return Loss":
        if row != col:
            return np.full_like(mag, np.nan, dtype=float), "Return Loss (dB)"
        return -20 * safe_log10(mag), "Return Loss (dB)"
    if data_type == "Insertion Loss":
        if row == col:
            return np.full_like(mag, np.nan, dtype=float), "Insertion Loss (dB)"
        return -20 * safe_log10(mag), "Insertion Loss (dB)"
    return 20 * safe_log10(mag), "Magnitude (dB)"


def flatten_sparams_map(sparams_map):
    all_sparams = []
    for sparams in sparams_map.values():
        all_sparams.extend(sparams)
    return sorted(set(all_sparams), key=sort_sparam_key)


def build_combined_dataframe(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale):
    frames = []
    for file_name in selected_files:
        network = network_dict[file_name]
        freq = network.f / freq_unit_scale
        for sparam in file_sparams_map.get(file_name, []):
            row, col = parse_sparameter(sparam)
            if row >= network.nports or col >= network.nports:
                continue
            y, _ = calculate_trace_data(network, sparam, data_type)
            frames.append(pd.DataFrame({
                "File": file_name,
                "Frequency": freq,
                "S-Parameter": sparam,
                "Data Type": data_type,
                "Value": y,
            }))
    if not frames:
        return pd.DataFrame(columns=["File", "Frequency", "S-Parameter", "Data Type", "Value"])
    return pd.concat(frames, ignore_index=True)


def build_marker_dataframe(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale, markers):
    rows = []
    for marker_idx, marker in enumerate(markers, start=1):
        for file_name in selected_files:
            network = network_dict[file_name]
            freq = network.f / freq_unit_scale
            if len(freq) == 0:
                continue
            nearest_idx = int(np.nanargmin(np.abs(freq - marker)))
            actual_freq = float(freq[nearest_idx])
            for sparam in file_sparams_map.get(file_name, []):
                row, col = parse_sparameter(sparam)
                if row >= network.nports or col >= network.nports:
                    continue
                y, _ = calculate_trace_data(network, sparam, data_type)
                value = y[nearest_idx]
                rows.append({
                    "Marker": f"M{marker_idx}",
                    "Target Frequency": marker,
                    "Actual Frequency": actual_freq,
                    "Frequency Delta": actual_freq - marker,
                    "File": file_name,
                    "S-Parameter": sparam,
                    "Data Type": data_type,
                    "Value": float(value) if not np.isnan(value) else np.nan,
                })
    return pd.DataFrame(rows)


def build_marker_display_table(marker_df, round_digits=2):
    if marker_df.empty:
        return pd.DataFrame()
    display_df = marker_df.pivot_table(
        index=["Marker", "Target Frequency", "File", "Data Type"],
        columns="S-Parameter",
        values="Value",
        aggfunc="first",
    ).reset_index()
    sparam_cols = [col for col in display_df.columns if re.fullmatch(r"S\d+\d+", str(col))]
    sparam_cols = sorted(sparam_cols, key=sort_sparam_key)
    display_df = display_df[["Marker", "Target Frequency", "File"] + sparam_cols + ["Data Type"]]
    display_df = display_df.rename(columns={"Data Type": "Datatype"})
    for col in sparam_cols:
        display_df[col] = display_df[col].round(round_digits)
    display_df["_MarkerSort"] = display_df["Marker"].str.extract(r"M(\d+)").astype(int)
    return display_df.sort_values(["_MarkerSort", "File"]).drop(columns=["_MarkerSort"])

# =========================================================
# Smith Chart Functions
# =========================================================
def add_smith_grid(fig, axis_limit=1.25):
    theta = np.linspace(0, 2 * np.pi, 600)
    fig.add_trace(go.Scatter(x=np.cos(theta), y=np.sin(theta), mode="lines", showlegend=False, hoverinfo="skip", line=dict(color="black", width=2)))
    fig.add_trace(go.Scatter(x=[-1, 1], y=[0, 0], mode="lines", showlegend=False, hoverinfo="skip", line=dict(color="lightgray", width=1)))

    r_values = [0.2, 0.5, 1.0, 2.0, 5.0]
    x_sweep = np.linspace(-50, 50, 1600)
    for r in r_values:
        z = r + 1j * x_sweep
        g = gamma_from_normalized_impedance(z)
        mask = np.abs(g) <= 1.0001
        fig.add_trace(go.Scatter(x=np.real(g[mask]), y=np.imag(g[mask]), mode="lines", showlegend=False, hoverinfo="skip", line=dict(color="gray", width=1.2)))
        g_label = gamma_from_normalized_impedance(r + 0j)
        fig.add_annotation(x=float(np.real(g_label)), y=0.025, text=f"{r:g}", showarrow=False, font=dict(size=12))

    x_values = [0.2, 0.5, 1.0, 2.0, 5.0]
    r_sweep = np.linspace(0, 80, 1600)
    for xval in x_values:
        for sign in [1, -1]:
            x = sign * xval
            z = r_sweep + 1j * x
            g = gamma_from_normalized_impedance(z)
            mask = np.abs(g) <= 1.0001
            fig.add_trace(go.Scatter(x=np.real(g[mask]), y=np.imag(g[mask]), mode="lines", showlegend=False, hoverinfo="skip", line=dict(color="gray" if xval != 1.0 else "black", width=1.1 if xval != 1.0 else 1.4)))
            g_label = gamma_from_normalized_impedance(0 + 1j * x)
            label = f"{xval:g}.0j" if xval >= 1 else f"{xval:g}j"
            if sign < 0:
                label = "-" + label
            fig.add_annotation(x=float(np.real(g_label)) * 1.06, y=float(np.imag(g_label)) * 1.06, text=label, showarrow=False, font=dict(size=12))

    fig.add_annotation(x=1.04, y=0, text="∞", showarrow=False, font=dict(size=18))
    fig.add_annotation(x=-1.05, y=0, text="0.0", showarrow=False, font=dict(size=12))
    fig.update_xaxes(range=[-axis_limit, axis_limit], visible=False, scaleanchor="y", scaleratio=1, constrain="domain")
    fig.update_yaxes(range=[-axis_limit, axis_limit], visible=False, constrain="domain")


def build_interactive_smith_chart(network_dict, selected_files, file_smith_sparams_map, freq_unit_scale, freq_unit, filter_start, filter_stop, markers, show_marker_points=True, smith_height=650, normalize_to_50=True, smith_target_z0=50.0, smith_full_range=False, smith_axis_limit=1.25):
    fig = go.Figure()
    add_smith_grid(fig, axis_limit=smith_axis_limit)

    for file_name in selected_files:
        selected_sparams = file_smith_sparams_map.get(file_name, [])
        if not selected_sparams:
            continue

        network = network_dict[file_name].copy()
        z0_note = "Original Z0"
        if normalize_to_50:
            try:
                network.renormalize(smith_target_z0)
                z0_note = f"Z0={smith_target_z0:g} ohm"
            except Exception:
                z0_note = "Original Z0, renormalize failed"

        freq = network.f / freq_unit_scale
        freq_mask = np.ones_like(freq, dtype=bool) if smith_full_range else (freq >= filter_start) & (freq <= filter_stop)

        for sparam in selected_sparams:
            row, col = parse_sparameter(sparam)
            if row >= network.nports or col >= network.nports:
                continue
            gamma = network.s[:, row, col]
            gamma_sel = gamma[freq_mask]
            freq_sel = freq[freq_mask]
            if len(gamma_sel) == 0:
                continue

            z_norm = normalized_impedance_from_gamma(gamma_sel)
            z_ohm = z_norm * smith_target_z0
            y_norm = normalized_admittance_from_z(z_norm)
            customdata = np.column_stack([
                freq_sel, np.real(gamma_sel), np.imag(gamma_sel), np.abs(gamma_sel), np.angle(gamma_sel, deg=True),
                np.real(z_norm), np.imag(z_norm), np.real(z_ohm), np.imag(z_ohm), np.real(y_norm), np.imag(y_norm),
            ])

            fig.add_trace(go.Scatter(
                x=np.real(gamma_sel),
                y=np.imag(gamma_sel),
                mode="lines",
                name=f"{file_name} - {sparam} ({z0_note})",
                customdata=customdata,
                line=dict(width=2.2),
                hovertemplate=(
                    f"<b>{file_name}</b><br>"
                    f"S-Parameter: {sparam}<br>"
                    f"Frequency: %{{customdata[0]:.6g}} {freq_unit}<br>"
                    "Γ Real: %{customdata[1]:.6g}<br>"
                    "Γ Imag: %{customdata[2]:.6g}<br>"
                    "|Γ|: %{customdata[3]:.6g}<br>"
                    "Phase: %{customdata[4]:.3f} deg<br>"
                    "z: %{customdata[5]:.3f} + j%{customdata[6]:.3f}<br>"
                    "Z: %{customdata[7]:.3f} + j%{customdata[8]:.3f} ohm<br>"
                    "y: %{customdata[9]:.3f} + j%{customdata[10]:.3f}<br>"
                    f"{z0_note}<extra></extra>"
                ),
            ))

            if show_marker_points and markers:
                mx, my, mtext, mcustom = [], [], [], []
                for marker_idx, marker in enumerate(markers, start=1):
                    idx = int(np.nanargmin(np.abs(freq - marker)))
                    actual_freq = float(freq[idx])
                    if smith_full_range or (filter_start <= actual_freq <= filter_stop):
                        g = gamma[idx]
                        z_m = normalized_impedance_from_gamma(g)
                        z_ohm_m = z_m * smith_target_z0
                        y_m = normalized_admittance_from_z(z_m)
                        mx.append(float(np.real(g)))
                        my.append(float(np.imag(g)))
                        mtext.append(f"M{marker_idx}")
                        mcustom.append([
                            marker, actual_freq, float(np.real(g)), float(np.imag(g)), float(np.abs(g)), float(np.angle(g, deg=True)),
                            float(np.real(z_m)), float(np.imag(z_m)), float(np.real(z_ohm_m)), float(np.imag(z_ohm_m)),
                            float(np.real(y_m)), float(np.imag(y_m)),
                        ])
                if mx:
                    fig.add_trace(go.Scatter(
                        x=mx,
                        y=my,
                        mode="markers+text",
                        showlegend=False,
                        marker=dict(size=12, symbol="x", line=dict(width=3)),
                        text=mtext,
                        textposition="top center",
                        customdata=np.array(mcustom),
                        hovertemplate=(
                            "<b>%{text}</b><br>"
                            f"{file_name} - {sparam}<br>"
                            f"Target: %{{customdata[0]:.6g}} {freq_unit}<br>"
                            f"Actual: %{{customdata[1]:.6g}} {freq_unit}<br>"
                            "Γ Real: %{customdata[2]:.6g}<br>"
                            "Γ Imag: %{customdata[3]:.6g}<br>"
                            "|Γ|: %{customdata[4]:.6g}<br>"
                            "Phase: %{customdata[5]:.3f} deg<br>"
                            "z: %{customdata[6]:.3f} + j%{customdata[7]:.3f}<br>"
                            "Z: %{customdata[8]:.3f} + j%{customdata[9]:.3f} ohm<br>"
                            "y: %{customdata[10]:.3f} + j%{customdata[11]:.3f}<extra></extra>"
                        ),
                    ))

    title = f"Smith Chart - normalized to {smith_target_z0:g} ohm" if normalize_to_50 else "Smith Chart"
    if smith_full_range:
        title += " - full frequency range"
    fig.update_layout(
        title=dict(text=title, x=0.5, xanchor="center", font=dict(size=24)),
        height=smith_height,
        template="plotly_white",
        hovermode="closest",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        margin=dict(l=20, r=20, t=90, b=20),
        plot_bgcolor="white",
    )
    return fig


def build_smith_marker_dataframe(network_dict, selected_files, file_smith_sparams_map, freq_unit_scale, markers, normalize_to_50=True, smith_target_z0=50.0):
    rows = []
    if not markers:
        return pd.DataFrame()
    for marker_idx, marker in enumerate(markers, start=1):
        for file_name in selected_files:
            selected_sparams = file_smith_sparams_map.get(file_name, [])
            if not selected_sparams:
                continue
            network = network_dict[file_name].copy()
            z0_note = "Original Z0"
            if normalize_to_50:
                try:
                    network.renormalize(smith_target_z0)
                    z0_note = f"Z0={smith_target_z0:g} ohm"
                except Exception:
                    z0_note = "Original Z0, renormalize failed"

            freq = network.f / freq_unit_scale
            idx = int(np.nanargmin(np.abs(freq - marker)))
            actual_freq = float(freq[idx])
            for sparam in selected_sparams:
                row, col = parse_sparameter(sparam)
                if row >= network.nports or col >= network.nports:
                    continue
                g = network.s[:, row, col][idx]
                z_norm = normalized_impedance_from_gamma(g)
                z_ohm = z_norm * smith_target_z0
                y_norm = normalized_admittance_from_z(z_norm)
                rows.append({
                    "Marker": f"M{marker_idx}",
                    "Target Frequency": marker,
                    "Actual Frequency": actual_freq,
                    "Frequency Delta": actual_freq - marker,
                    "File": file_name,
                    "S-Parameter": sparam,
                    "Gamma Real": float(np.real(g)),
                    "Gamma Imag": float(np.imag(g)),
                    "Gamma Magnitude": float(np.abs(g)),
                    "Gamma Phase deg": float(np.angle(g, deg=True)),
                    "z Real": float(np.real(z_norm)),
                    "z Imag": float(np.imag(z_norm)),
                    "Z Real ohm": float(np.real(z_ohm)),
                    "Z Imag ohm": float(np.imag(z_ohm)),
                    "y Real": float(np.real(y_norm)),
                    "y Imag": float(np.imag(y_norm)),
                    "Z0": z0_note,
                })
    return pd.DataFrame(rows)


def make_excel_bytes(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale, filter_start, filter_stop, marker_df, summary_df, smith_marker_df=None):
    output = BytesIO()
    used_sheet_names = set()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        used_sheet_names.add("Summary")

        if not marker_df.empty:
            marker_df.to_excel(writer, sheet_name="Marker Raw", index=False)
            used_sheet_names.add("Marker Raw")
            marker_display_df = build_marker_display_table(marker_df)
            marker_display_df.to_excel(writer, sheet_name="Marker Compare", index=False)
            used_sheet_names.add("Marker Compare")

        if smith_marker_df is not None and not smith_marker_df.empty:
            smith_marker_df.to_excel(writer, sheet_name="Smith Markers", index=False)
            used_sheet_names.add("Smith Markers")

        for file_name in selected_files:
            network = network_dict[file_name]
            freq = network.f / freq_unit_scale
            freq_mask = (freq >= filter_start) & (freq <= filter_stop)
            file_df = pd.DataFrame({"Frequency": freq[freq_mask]})
            for sparam in file_sparams_map.get(file_name, []):
                row, col = parse_sparameter(sparam)
                if row >= network.nports or col >= network.nports:
                    continue
                y, _ = calculate_trace_data(network, sparam, data_type)
                file_df[f"{sparam}_{data_type}"] = y[freq_mask]
            sheet_name = make_unique_sheet_name(file_name, used_sheet_names)
            file_df.to_excel(writer, sheet_name=sheet_name, index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            if worksheet.max_row >= 1 and worksheet.max_column >= 1:
                worksheet.auto_filter.ref = worksheet.dimensions
            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter
                for cell in column_cells:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)
    return output.getvalue()

# =========================================================
# Floating toolbar UI
# =========================================================
st.markdown("<span class='small-note'>Floating-style tool panels：Files / Layout / Markers / Smith / Traces / Range / Export</span>", unsafe_allow_html=True)

top_spacer, file_col, layout_col, marker_col, smith_col = st.columns([5.0, 1.05, 1.05, 1.05, 1.05])
with top_spacer:
    st.markdown("<span class='tool-title'>SNP Viewer Tool Palette</span>", unsafe_allow_html=True)

with file_col:
    with open_panel("📁 Files", key="panel_files"):
        uploaded_files = st.file_uploader(
            "Upload Touchstone Files",
            type=["s1p", "s2p", "s3p", "s4p", "s5p", "s6p", "s7p", "s8p", "snp", "txt"],
            accept_multiple_files=True,
        )

with layout_col:
    with open_panel("⚙️ Layout", key="panel_layout"):
        freq_unit = st.selectbox("Frequency Unit", ["Hz", "kHz", "MHz", "GHz"], index=2)
        freq_unit_scale = {"Hz": 1, "kHz": 1e3, "MHz": 1e6, "GHz": 1e9}[freq_unit]
        data_type = st.selectbox("Data Type", ["Magnitude dB", "Phase deg", "VSWR", "Return Loss", "Insertion Loss"], index=0)
        line_mode = st.selectbox("Plot Mode", ["lines", "lines+markers", "markers"], index=0)
        plot_height = st.slider("Plot Height", 400, 1000, 700, 50)
        st.divider()
        y_axis_manual = st.checkbox("Manual Y Axis Range", value=False)
        y_axis_min_input = st.number_input("Y Axis Min", value=-40.0, step=1.0, disabled=not y_axis_manual)
        y_axis_max_input = st.number_input("Y Axis Max", value=0.0, step=1.0, disabled=not y_axis_manual)

with marker_col:
    with open_panel("📍 Markers", key="panel_markers"):
        st.text_area(f"Marker Frequencies ({freq_unit})", key="marker_text_buffer", placeholder="Example: 2400, 2450, 2500")
        st.button("Update Markers", type="primary", use_container_width=True, on_click=update_markers_from_input)
        if st.session_state.applied_markers:
            st.success("Applied: " + ", ".join([f"{m:g}" for m in st.session_state.applied_markers]))
        else:
            st.info("Input Marker and click Update Markers.")
        show_marker_lines = st.checkbox("Show Marker Vertical Lines", value=True)
        show_marker_points = st.checkbox("Show Marker Points", value=True)
        markers = st.session_state.applied_markers

with smith_col:
    with open_panel("🧭 Smith", key="panel_smith"):
        show_smith_chart = st.checkbox("Show Smith Chart", value=True)
        smith_freq_mode = st.selectbox("Smith Chart Frequency Range", ["Full Range", "Same as Main Plot", "Custom Range"], index=0)
        smith_normalize_to_50 = st.checkbox("Normalize Smith Chart to 50 ohm", value=True)
        smith_target_z0 = st.number_input("Smith Chart Reference Z0 (ohm)", min_value=1.0, max_value=1000.0, value=50.0, step=1.0)
        smith_height = st.slider("Smith Chart Height", 400, 900, 650, 50)

if not uploaded_files:
    st.info("請從上方 Files 面板上傳一個或多個 S1P ~ S8P 檔案。")
    st.stop()

# =========================================================
# Read files
# =========================================================
network_dict = {}
file_info = []
for uploaded_file in uploaded_files:
    try:
        network = read_touchstone(uploaded_file)
        network_dict[uploaded_file.name] = network
        file_info.append({
            "File": uploaded_file.name,
            "Ports": network.nports,
            "Frequency Points": len(network.f),
            f"Start Frequency ({freq_unit})": network.f[0] / freq_unit_scale,
            f"Stop Frequency ({freq_unit})": network.f[-1] / freq_unit_scale,
            "Z0 first point": str(network.z0[0] if hasattr(network, "z0") else "N/A"),
        })
    except Exception as e:
        st.error(f"讀取失敗：{uploaded_file.name}")
        st.exception(e)

if not network_dict:
    st.error("沒有任何檔案成功讀取。")
    st.stop()

st.success(f"成功讀取 {len(network_dict)} 個檔案")
with st.expander("File Information", expanded=False):
    st.dataframe(pd.DataFrame(file_info), use_container_width=True)

# =========================================================
# Second toolbar row: content-dependent settings
# =========================================================
all_file_names = list(network_dict.keys())
trace_spacer, trace_col, range_col, export_col, help_col = st.columns([5.6, 1.05, 1.05, 1.05, 1.05])
with trace_spacer:
    st.markdown("<span class='small-note'>Trace 選項已依 SNP 檔案分開：每個檔案可以選不同的 Main Plot 與 Smith Chart S-Parameters。</span>", unsafe_allow_html=True)

with trace_col:
    with open_panel("📈 Traces", key="panel_traces"):
        selected_files = st.multiselect("Select Files to Compare", options=all_file_names, default=all_file_names)
        if not selected_files:
            st.warning("請至少選擇一個檔案。")
            st.stop()

        st.divider()
        st.markdown("**Main Plot Trace per SNP file**")
        file_sparams_map = {}
        file_smith_sparams_map = {}

        for file_name in selected_files:
            network = network_dict[file_name]
            available_sparams = get_sparameter_list(network.nports)
            default_sparams = [s for s in ["S11", "S21"] if s in available_sparams]
            if not default_sparams and available_sparams:
                default_sparams = [available_sparams[0]]

            with st.expander(f"Trace: {file_name}", expanded=True):
                file_sparams_map[file_name] = st.multiselect(
                    "Select S-Parameters",
                    options=available_sparams,
                    default=default_sparams,
                    key=f"trace_sparams_{file_name}",
                )

                reflection_sparams = get_reflection_sparams(available_sparams)
                smith_defaults = [s for s in ["S11"] if s in reflection_sparams]
                if not smith_defaults and reflection_sparams:
                    smith_defaults = [reflection_sparams[0]]

                if show_smith_chart:
                    file_smith_sparams_map[file_name] = st.multiselect(
                        "Select Smith Chart S-Parameters",
                        options=reflection_sparams,
                        default=smith_defaults,
                        key=f"smith_sparams_{file_name}",
                    )
                else:
                    file_smith_sparams_map[file_name] = []

        if not any(file_sparams_map.get(file_name) for file_name in selected_files):
            st.warning("請至少在一個檔案選擇一個 Main Plot S-Parameter。")
            st.stop()

# Fallback defaults if Streamlit rerun evaluates below before popover variables exist.
if "selected_files" not in locals():
    selected_files = all_file_names
if "file_sparams_map" not in locals():
    file_sparams_map = {}
    for file_name in selected_files:
        available_sparams = get_sparameter_list(network_dict[file_name].nports)
        file_sparams_map[file_name] = [s for s in ["S11", "S21"] if s in available_sparams] or available_sparams[:1]
if "file_smith_sparams_map" not in locals():
    file_smith_sparams_map = {}
    for file_name in selected_files:
        reflection_sparams = get_reflection_sparams(get_sparameter_list(network_dict[file_name].nports))
        file_smith_sparams_map[file_name] = (["S11"] if "S11" in reflection_sparams else reflection_sparams[:1]) if show_smith_chart else []

all_freq_values = np.concatenate([network_dict[file].f / freq_unit_scale for file in selected_files])
freq_min = float(np.nanmin(all_freq_values))
freq_max = float(np.nanmax(all_freq_values))

with range_col:
    with open_panel("🔎 Range", key="panel_range"):
        filter_start = st.number_input(f"Start Frequency ({freq_unit})", value=freq_min, min_value=freq_min, max_value=freq_max)
        filter_stop = st.number_input(f"Stop Frequency ({freq_unit})", value=freq_max, min_value=freq_min, max_value=freq_max)
        if filter_start >= filter_stop:
            st.error("Start Frequency 必須小於 Stop Frequency。")
            st.stop()
        st.divider()
        st.caption("Smith Chart range mode is selected in the Smith panel. Custom start/stop appears below the Smith Chart.")

if "filter_start" not in locals():
    filter_start = freq_min
if "filter_stop" not in locals():
    filter_stop = freq_max
if filter_start >= filter_stop:
    st.error("Start Frequency 必須小於 Stop Frequency。")
    st.stop()

if st.session_state.smith_filter_start == 0.0 and st.session_state.smith_filter_stop == 0.0:
    st.session_state.smith_filter_start = freq_min
    st.session_state.smith_filter_stop = freq_max

if show_smith_chart:
    if smith_freq_mode == "Full Range":
        smith_full_range = True
        smith_filter_start = freq_min
        smith_filter_stop = freq_max
    elif smith_freq_mode == "Same as Main Plot":
        smith_full_range = False
        smith_filter_start = filter_start
        smith_filter_stop = filter_stop
    else:
        smith_full_range = False
        smith_filter_start = float(st.session_state.smith_filter_start)
        smith_filter_stop = float(st.session_state.smith_filter_stop)
    if smith_filter_start >= smith_filter_stop:
        st.error("Smith Start Frequency must be smaller than Smith Stop Frequency.")
        st.stop()
else:
    smith_full_range = True
    smith_filter_start = freq_min
    smith_filter_stop = freq_max

with export_col:
    export_panel_slot = st.empty()
    with export_panel_slot.container():
        with open_panel("⬇️ Export", key="panel_export_pending"):
            st.caption("準備下載資料中；檔案讀取完成後會自動啟用。")
            st.markdown("Download CSV")
            st.markdown("Download Excel")
            st.markdown("Download HTML")
            st.markdown("Download PNG")
            st.markdown("Download Smith PNG")

with help_col:
    with open_panel("❓ Help", key="panel_help"):
        st.markdown(
            """
            **Popover UI 說明**
            - **Files**：上傳 Touchstone 檔案
            - **Layout**：資料格式、圖高、Y 軸範圍
            - **Traces**：先選檔案，再針對每個 SNP 檔案獨立選 Main Plot 與 Smith Chart S-Parameter
            - **Markers**：輸入 Marker 並更新
            - **Smith**：Smith Chart 設定
            - **Range**：主圖與 Smith Chart 頻率範圍
            - **Export**：下載 CSV / Excel / HTML / PNG
            """
        )

# =========================================================
# Marker table
# =========================================================
marker_df = build_marker_dataframe(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale, st.session_state.applied_markers)
markers = st.session_state.applied_markers
if markers:
    st.subheader("Marker Value Table")
    st.dataframe(build_marker_display_table(marker_df), use_container_width=True)
    with st.expander("Marker Raw Data", expanded=False):
        st.dataframe(marker_df, use_container_width=True)

# =========================================================
# Main plot
# =========================================================
fig = go.Figure()
ylabel = "Value"
for file_name in selected_files:
    network = network_dict[file_name]
    freq = network.f / freq_unit_scale
    freq_mask = (freq >= filter_start) & (freq <= filter_stop)

    for sparam in file_sparams_map.get(file_name, []):
        row, col = parse_sparameter(sparam)
        if row >= network.nports or col >= network.nports:
            continue
        y, ylabel = calculate_trace_data(network, sparam, data_type)
        fig.add_trace(go.Scatter(
            x=freq[freq_mask],
            y=y[freq_mask],
            mode=line_mode,
            name=f"{file_name} - {sparam}",
            hovertemplate=(
                f"<b>{file_name}</b><br>"
                f"S-Parameter: {sparam}<br>"
                f"Frequency: %{{x:.6g}} {freq_unit}<br>"
                f"{data_type}: %{{y:.6g}}<extra></extra>"
            ),
        ))

        if show_marker_points and markers:
            marker_x, marker_y, marker_label, marker_hover = [], [], [], []
            for marker_idx, marker in enumerate(markers, start=1):
                idx = int(np.nanargmin(np.abs(freq - marker)))
                actual_freq = float(freq[idx])
                if filter_start <= actual_freq <= filter_stop:
                    marker_x.append(actual_freq)
                    marker_y.append(y[idx])
                    marker_label.append(f"M{marker_idx}")
                    marker_hover.append(
                        f"Marker: M{marker_idx}<br>"
                        f"Target: {marker:.6g} {freq_unit}<br>"
                        f"Actual: {actual_freq:.6g} {freq_unit}<br>"
                        f"{file_name} - {sparam}<br>"
                        f"{data_type}: {y[idx]:.6g}"
                    )
            if marker_x:
                fig.add_trace(go.Scatter(
                    x=marker_x,
                    y=marker_y,
                    mode="markers+text",
                    showlegend=False,
                    marker=dict(size=11, symbol="x", line=dict(width=2)),
                    text=marker_label,
                    textposition="top center",
                    hovertext=marker_hover,
                    hovertemplate="%{hovertext}<extra></extra>",
                ))

if show_marker_lines and markers:
    for marker_idx, marker in enumerate(markers, start=1):
        if filter_start <= marker <= filter_stop:
            fig.add_vline(
                x=marker,
                line_width=1,
                line_dash="dash",
                line_color="gray",
                annotation_text=f"M{marker_idx}: {marker:g} {freq_unit}",
                annotation_position="top",
            )

fig.update_layout(
    title=f"SNP Compare - {data_type}",
    xaxis_title=f"Frequency ({freq_unit})",
    yaxis_title=ylabel,
    hovermode="x unified",
    height=plot_height,
    legend_title="Trace",
    template="plotly_white",
)
fig.update_xaxes(rangeslider_visible=True)

if y_axis_manual:
    if y_axis_min_input >= y_axis_max_input:
        st.warning("Y Axis Min must be smaller than Y Axis Max. Auto Y-axis is used now.")
    else:
        fig.update_yaxes(range=[y_axis_min_input, y_axis_max_input])

st.plotly_chart(fig, use_container_width=True)

# =========================================================
# Smith Chart
# =========================================================
if show_smith_chart and any(file_smith_sparams_map.get(f) for f in selected_files):
    smith_fig = build_interactive_smith_chart(
        network_dict=network_dict,
        selected_files=selected_files,
        file_smith_sparams_map=file_smith_sparams_map,
        freq_unit_scale=freq_unit_scale,
        freq_unit=freq_unit,
        filter_start=smith_filter_start,
        filter_stop=smith_filter_stop,
        markers=markers,
        show_marker_points=show_marker_points,
        smith_height=smith_height,
        normalize_to_50=smith_normalize_to_50,
        smith_target_z0=smith_target_z0,
        smith_full_range=smith_full_range,
        smith_axis_limit=float(st.session_state.smith_axis_limit),
    )
    smith_width_pct = int(st.session_state.smith_display_width_pct)
    side_pct = max((100 - smith_width_pct) / 2, 1)
    if smith_width_pct >= 98:
        st.plotly_chart(smith_fig, use_container_width=True)
    else:
        smith_left, smith_center, smith_right = st.columns([side_pct, smith_width_pct, side_pct])
        with smith_center:
            st.plotly_chart(smith_fig, use_container_width=True)

    if markers and show_marker_points:
        smith_marker_df = build_smith_marker_dataframe(
            network_dict,
            selected_files,
            file_smith_sparams_map,
            freq_unit_scale,
            markers,
            smith_normalize_to_50,
            smith_target_z0,
        )
        with st.expander("Smith Chart Marker Table", expanded=False):
            st.dataframe(smith_marker_df, use_container_width=True)
    else:
        smith_marker_df = pd.DataFrame()

    with st.expander("Smith Chart Display / Frequency Controls", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            st.slider("Smith Chart Display Width (%)", min_value=40, max_value=100, value=int(st.session_state.smith_display_width_pct), step=5, key="smith_display_width_pct")
        with c2:
            st.slider("Smith Chart Axis Limit", min_value=1.05, max_value=2.00, value=float(st.session_state.smith_axis_limit), step=0.05, key="smith_axis_limit")

        if smith_freq_mode == "Full Range":
            st.info(f"Smith Chart Frequency Range: Full Range ({freq_min:g} ~ {freq_max:g} {freq_unit})")
        elif smith_freq_mode == "Same as Main Plot":
            st.info(f"Smith Chart Frequency Range: Same as Main Plot ({filter_start:g} ~ {filter_stop:g} {freq_unit})")
        else:
            st.subheader("Smith Chart Custom Frequency Range")
            smith_c1, smith_c2 = st.columns(2)
            with smith_c1:
                st.number_input(f"Smith Start Frequency ({freq_unit})", value=float(st.session_state.smith_filter_start), min_value=freq_min, max_value=freq_max, key="smith_filter_start")
            with smith_c2:
                st.number_input(f"Smith Stop Frequency ({freq_unit})", value=float(st.session_state.smith_filter_stop), min_value=freq_min, max_value=freq_max, key="smith_filter_stop")
else:
    smith_fig = None
    smith_marker_df = pd.DataFrame()

# =========================================================
# Warnings
# =========================================================
selected_sparams_flat = flatten_sparams_map(file_sparams_map)
if data_type in ["VSWR", "Return Loss"]:
    invalid = [s for s in selected_sparams_flat if parse_sparameter(s)[0] != parse_sparameter(s)[1]]
    if invalid:
        st.warning(f"{data_type} usually applies to reflection parameters such as S11/S22. These will be NaN: {', '.join(invalid)}")
if data_type == "Insertion Loss":
    invalid = [s for s in selected_sparams_flat if parse_sparameter(s)[0] == parse_sparameter(s)[1]]
    if invalid:
        st.warning(f"Insertion Loss usually applies to transmission parameters such as S21/S12. These will be NaN: {', '.join(invalid)}")

# =========================================================
# Export data preparation
# =========================================================
combined_df = build_combined_dataframe(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale)
combined_df = combined_df[(combined_df["Frequency"] >= filter_start) & (combined_df["Frequency"] <= filter_stop)]

with st.expander("Preview Export Data", expanded=False):
    st.dataframe(combined_df, use_container_width=True)

main_selection_text = "; ".join([f"{f}: {', '.join(file_sparams_map.get(f, []))}" for f in selected_files])
smith_selection_text = "; ".join([f"{f}: {', '.join(file_smith_sparams_map.get(f, []))}" for f in selected_files])
summary_df = pd.DataFrame({
    "Item": [
        "Export Date", "Frequency Unit", "Data Type", "Selected Files", "Selected Main Plot S-Parameters per File",
        "Selected Smith S-Parameters per File", "Filter Start", "Filter Stop", "Y Axis Manual", "Y Axis Min", "Y Axis Max",
        "Applied Markers", "Smith Normalize", "Smith Z0", "Smith Frequency Mode", "Smith Start", "Smith Stop",
    ],
    "Value": [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), freq_unit, data_type, ", ".join(selected_files), main_selection_text,
        smith_selection_text, filter_start, filter_stop, y_axis_manual, y_axis_min_input, y_axis_max_input,
        ", ".join([str(m) for m in markers]) if markers else "", smith_normalize_to_50, smith_target_z0,
        smith_freq_mode, smith_filter_start, smith_filter_stop,
    ],
})

csv_data = combined_df.to_csv(index=False).encode("utf-8-sig")
try:
    excel_data = make_excel_bytes(network_dict, selected_files, file_sparams_map, data_type, freq_unit_scale, filter_start, filter_stop, marker_df, summary_df, smith_marker_df=smith_marker_df)
except Exception as e:
    excel_data = None
    st.error("Excel export failed. Please install openpyxl: pip install openpyxl")
    st.exception(e)

png_data = None
smith_png_data = None
if st.session_state.enable_png_export:
    try:
        png_data = fig.to_image(format="png", scale=2)
    except Exception:
        png_data = None
    try:
        smith_png_data = smith_fig.to_image(format="png", scale=2) if smith_fig is not None else None
    except Exception:
        smith_png_data = None

html_text = fig.to_html(include_plotlyjs="cdn", full_html=False)
if smith_fig is not None:
    html_text += "<br><hr><br>" + smith_fig.to_html(include_plotlyjs=False, full_html=False)
html_data = html_text.encode("utf-8")

export_panel_slot.empty()
with export_panel_slot.container():
    with open_panel("✅ Export", key="panel_export_active"):
        st.markdown("✅ 資料已準備完成，可以立即下載 CSV / Excel / HTML。")
        st.checkbox("Generate PNG exports (slower)", key="enable_png_export")
        if not st.session_state.enable_png_export:
            st.caption("為了加速讀取，PNG 預設不預先產生；勾選後會重新執行並產生 PNG。")

        st.download_button("✅ Download CSV", data=csv_data, file_name=f"snp_compare_result_{TODAY_MMDD}.csv", mime="text/csv", use_container_width=True, key="download_csv_active")
        if excel_data is not None:
            st.download_button("✅ Download Excel", data=excel_data, file_name=f"snp_compare_result_{TODAY_MMDD}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="download_excel_active")
        else:
            st.button("Download Excel", disabled=True, use_container_width=True, key="download_excel_disabled_active")

        st.download_button("✅ Download HTML", data=html_data, file_name=f"snp_compare_plot_{TODAY_MMDD}.html", mime="text/html", use_container_width=True, key="download_html_active")

        if png_data is not None:
            st.download_button("✅ Download PNG", data=png_data, file_name=f"snp_compare_plot_{TODAY_MMDD}.png", mime="image/png", use_container_width=True, key="download_png_active")
        else:
            st.button("Download PNG", disabled=True, use_container_width=True, key="download_png_disabled_active")
            st.caption("PNG 未產生：勾選 Generate PNG exports，並確認已安裝 kaleido。")

        if smith_png_data is not None:
            st.download_button("✅ Download Smith PNG", data=smith_png_data, file_name=f"smith_chart_{TODAY_MMDD}.png", mime="image/png", use_container_width=True, key="download_smith_png_active")
        else:
            st.button("Smith PNG", disabled=True, use_container_width=True, key="download_smith_png_disabled_active")
            st.caption("Smith PNG 未產生：勾選 Generate PNG exports，並確認 Smith Chart 已啟用。")

# =========================================================
# Notes
# =========================================================
with st.expander("Calculation Notes"):
    st.markdown(
        """
        ### Smith Chart hover values
        Smith Chart line hover includes `z`, `Z`, and `y`:
        ```text
        Γ = Sii
        z = (1 + Γ) / (1 - Γ)
        Z = z * Z0
        y = 1 / z
        ```
        - `z` is normalized impedance.
        - `Z` is actual impedance in ohm.
        - `y` is normalized admittance.

        ### Magnitude dB
        ```text
        Magnitude dB = 20 * log10(|Sij|)
        ```

        ### VSWR
        ```text
        VSWR = (1 + |Γ|) / (1 - |Γ|)
        ```

        ### Return Loss
        ```text
        Return Loss = -20 * log10(|Sii|)
        ```

        ### Insertion Loss
        ```text
        Insertion Loss = -20 * log10(|Sij|)
        ```
        """
    )
