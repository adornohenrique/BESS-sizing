import math
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# Utility functions
# ---------------------------------------------------------

def round_up_to_4_dec(x: float) -> float:
    """Mimic Excel ROUNDUP(...,4) behaviour."""
    if x is None:
        return None
    factor = 10_000
    return math.ceil(x * factor) / factor


def select_pcs_rating(container_power_mw: float) -> float:
    """
    Approximate the PCS selection logic from the Excel:
    choose the smallest rating >= required power.
    Available PCS sizes (MW): 1.25, 1.5, 1.75, 2, 2.5, 5
    """
    pcs_sizes = [1.25, 1.50, 1.75, 2.0, 2.5, 5.0]
    for s in pcs_sizes:
        if container_power_mw <= s:
            return s
    return pcs_sizes[-1]


def select_transformer_mva(container_power_mw: float) -> str:
    """
    Approximate transformer MVA selection similar to Excel I53/I55:
    smallest MVA >= container power.
    """
    candidates = [1.25, 1.5, 1.75, 2.0, 2.5, 3.5, 5.0]
    for c in candidates:
        if container_power_mw <= c:
            return f"{c:.2f}MVA".rstrip("0").rstrip(".") + "MVA"
    return f"{candidates[-1]:.2f}MVA".rstrip("0").rstrip(".") + "MVA"


def cable_runs_300mm(current_amps: float) -> str:
    """
    Reproduce the cable remark logic:
    300 mm² cable per run carrying up to ~446 A,
    then add runs as needed.
    """
    if current_amps <= 0:
        return "No current"
    limit_per_run = 446.0
    runs = math.ceil(current_amps / limit_per_run)
    if runs == 1:
        return "Use 300mm² cable, 1 run per phase"
    return f"Use 300mm² cable, {runs} runs per phase"


# ---------------------------------------------------------
# Streamlit layout
# ---------------------------------------------------------

st.set_page_config(
    page_title="BESS Sizing Calculator (Web)",
    layout="wide"
)

st.title("BESS Sizing Calculator (Web Version)")
st.caption("Converted from 'BESS Sizing Calculator (R2.1)' Excel sheet")

with st.sidebar:
    st.subheader("Site Load Requirements")

    load_mw = st.number_input(
        "Customer load supported by BESS (MW)",
        min_value=0.0,
        value=0.99,
        step=0.01
    )

    discharge_h = st.number_input(
        "Discharge duration (hours)",
        min_value=0.0,
        value=8.0,
        step=0.5
    )

    dod_percent = st.number_input(
        "Depth of Discharge, DoD (%)",
        min_value=1.0,
        max_value=100.0,
        value=90.0,
        step=1.0
    )

    rte_percent = st.number_input(
        "Round-trip efficiency, RTE (%)",
        min_value=1.0,
        max_value=100.0,
        value=88.65,
        step=0.1
    )

    other_eff_percent = st.number_input(
        "Other efficiency (%)",
        min_value=1.0,
        max_value=100.0,
        value=100.0,
        step=0.1
    )

    st.markdown("---")
    st.subheader("Operating Parameters")

    c_rate = st.number_input(
        "Customer C-rate (max 0.5C)",
        min_value=0.01,
        max_value=1.0,
        value=0.25,
        step=0.01
    )

    grid_charging_mw = st.number_input(
        "Power available for charging – grid (MW)",
        min_value=0.0,
        value=0.5,
        step=0.1
    )

    other_charging_mw = st.number_input(
        "Power available for charging – other (MW)",
        min_value=0.0,
        value=1.0,
        step=0.1
    )

    st.markdown("---")
    st.subheader("Electrical Data")

    voltage_kv = st.number_input(
        "Voltage standard (kV)",
        min_value=0.1,
        value=0.4,
        step=0.01
    )

    power_factor = st.number_input(
        "Power factor",
        min_value=0.1,
        max_value=1.0,
        value=0.85,
        step=0.01
    )

    st.markdown("---")
    st.subheader("Project Info (Qualitative)")

    project_application = st.selectbox(
        "Project application",
        [
            "Time of Use (TOU) Arbitrage",
            "Peak shaving",
            "Backup / Black start",
            "Frequency regulation",
            "Other"
        ],
        index=0
    )

    ambient_env = st.selectbox(
        "Ambient environment",
        ["Inland", "Coastal", "Harsh / Industrial"],
        index=0
    )

    cooling = st.selectbox(
        "Cooling system",
        ["Liquid cooling system", "Air cooling"],
        index=0
    )

    cycles_per_day = st.number_input(
        "Charge/discharge cycles per day",
        min_value=0.0,
        value=1.0,
        step=0.5
    )

    black_start = st.selectbox(
        "Black start capability",
        ["Not required", "Required"],
        index=0
    )

# ---------------------------------------------------------
# Core BESS sizing (mimicking Excel logic)
# ---------------------------------------------------------

# Step 1 – Required capacity
initial_mwh = load_mw * discharge_h  # I19 = E19*E20

after_dod_mwh = round_up_to_4_dec(initial_mwh / (dod_percent / 100.0))  # I20
after_rte_mwh = round_up_to_4_dec(after_dod_mwh / (rte_percent / 100.0))  # I21
after_other_eff_mwh = round_up_to_4_dec(after_rte_mwh / (other_eff_percent / 100.0))  # I22

required_bess_mwh = after_other_eff_mwh  # I27 ~ I22 (simplified: no E5/E4 special case)
required_discharge_power_mw = load_mw  # I25 = E19
customer_c_rate = c_rate               # I26 = E28
charging_power_mw = round(grid_charging_mw + other_charging_mw, 2)  # I32

# ---------------------------------------------------------
# Battery models (from original Excel sheet)
# ---------------------------------------------------------

battery_models = [
    {"name": "261 kWh battery", "capacity_kwh": 261.0},
    {"name": "3727.36 kWh battery", "capacity_kwh": 3727.36},
    {"name": "5015.9 kWh battery", "capacity_kwh": 5015.9},
]

rows = []

for model in battery_models:
    cap_kwh = model["capacity_kwh"]

    if required_bess_mwh > 0:
        units = math.ceil(required_bess_mwh * 1000.0 / cap_kwh)
    else:
        units = 0

    total_mwh = units * cap_kwh / 1000.0
    oversize_mwh = total_mwh - required_bess_mwh if required_bess_mwh is not None else None
    oversize_pct = (
        (oversize_mwh / required_bess_mwh * 100.0)
        if required_bess_mwh and required_bess_mwh > 0
        else 0.0
    )

    # PCS sizing per unit – similar idea to I36/I41/I38/I43
    container_power_mw = customer_c_rate * cap_kwh / 1000.0
    pcs_rating_mw = select_pcs_rating(container_power_mw)
    pcs_quantity = math.ceil(required_discharge_power_mw / pcs_rating_mw) if pcs_rating_mw > 0 else 0

    # Transformer sizing – similar idea to I53/I55
    transformer_mva = select_transformer_mva(container_power_mw)

    rows.append({
        "Battery model": model["name"],
        "Unit capacity (kWh)": cap_kwh,
        "Number of units": units,
        "Total installed capacity (MWh)": total_mwh,
        "Oversizing vs required (MWh)": oversize_mwh,
        "Oversizing (%)": oversize_pct,
        "Container C-rate (C)": customer_c_rate,
        "Approx. PCS rating per unit (MW)": pcs_rating_mw,
        "PCS quantity (for discharge power)": pcs_quantity,
        "Transformer per unit (approx)": transformer_mva,
    })

df_bess = pd.DataFrame(rows)

# Identify "best" model by smallest oversizing
best_idx = None
if required_bess_mwh and required_bess_mwh > 0:
    best_idx = df_bess["Oversizing vs required (MWh)"].clip(lower=0).idxmin()

# ---------------------------------------------------------
# Electrical – main breaker and cable
# ---------------------------------------------------------

if voltage_kv > 0 and power_factor > 0:
    current_amps = required_discharge_power_mw * 1_000_000.0 / (math.sqrt(3) * voltage_kv * 1000.0 * power_factor)
else:
    current_amps = 0.0

breaker_amps = current_amps * 1.25  # apply 125% NEC factor like I59

cable_recommendation = cable_runs_300mm(current_amps)

# ---------------------------------------------------------
# Display results
# ---------------------------------------------------------

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Required Battery Capacity (Excel Logic)")

    st.metric("Customer load (MW)", f"{load_mw:.3f}")
    st.metric("Discharge duration (h)", f"{discharge_h:.2f}")
    st.metric("Initial energy (MWh)", f"{initial_mwh:.4f}")

    st.write("**After applying efficiencies:**")
    st.write(f"- After DoD: **{after_dod_mwh:.4f} MWh**")
    st.write(f"- After RTE: **{after_rte_mwh:.4f} MWh**")
    st.write(f"- After other efficiency: **{after_other_eff_mwh:.4f} MWh**")

    st.markdown("---")
    st.metric("Required BESS capacity (MWh)", f"{required_bess_mwh:.4f}")
    st.metric("Required discharging power (MW)", f"{required_discharge_power_mw:.3f}")
    st.metric("Charging power available (MW)", f"{charging_power_mw:.2f}")

with col2:
    st.subheader("2. Electrical Summary")

    st.metric("System voltage (kV)", f"{voltage_kv:.3f}")
    st.metric("Power factor", f"{power_factor:.2f}")

    st.write(f"**Line current (approx.):** {current_amps:,.2f} A")
    st.write(f"**Main breaker rating (125%):** {breaker_amps:,.2f} A")
    st.write(f"**Cable recommendation:** {cable_recommendation}")

    st.markdown("---")
    st.write("**Qualitative info:**")
    st.write(f"- Application: {project_application}")
    st.write(f"- Environment: {ambient_env}")
    st.write(f"- Cooling: {cooling}")
    st.write(f"- Cycles per day: {cycles_per_day}")
    st.write(f"- Black start capability: {black_start}")

st.markdown("---")
st.subheader("3. BESS Configuration Options")

st.dataframe(
    df_bess.style.format({
        "Unit capacity (kWh)": "{:,.2f}",
        "Number of units": "{:,.0f}",
        "Total installed capacity (MWh)": "{:,.4f}",
        "Oversizing vs required (MWh)": "{:,.4f}",
        "Oversizing (%)": "{:,.2f}",
        "Container C-rate (C)": "{:,.2f}",
        "Approx. PCS rating per unit (MW)": "{:,.2f}",
        "PCS quantity (for discharge power)": "{:,.0f}",
    })
)

if best_idx is not None and 0 <= best_idx < len(df_bess):
    best_row = df_bess.iloc[best_idx]
    st.markdown("### 4. Suggested Configuration")
    st.write(
        f"- **Battery model:** {best_row['Battery model']}  \n"
        f"- **Number of units:** {int(best_row['Number of units'])}  \n"
        f"- **Total installed capacity:** {best_row['Total installed capacity (MWh)']:.4f} MWh  \n"
        f"- **Oversizing:** {best_row['Oversizing vs required (MWh)']:.4f} MWh "
        f"({best_row['Oversizing (%)']:.2f}%)  \n"
        f"- **PCS per unit:** ~{best_row['Approx. PCS rating per unit (MW)']:.2f} MW  \n"
        f"- **PCS quantity:** {int(best_row['PCS quantity (for discharge power)'])}  \n"
        f"- **Transformer per unit:** {best_row['Transformer per unit (approx)']}"
    )
