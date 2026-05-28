import re
from pathlib import Path

import pandas as pd
import streamlit as st
import folium
from folium.plugins import HeatMap
from streamlit_folium import st_folium
import pgeocode


# =========================================================
# 1. Online data location
# =========================================================

BASE_DIR = Path(__file__).resolve().parent / "data"
DATA_FILE = BASE_DIR / "HeatmapData.xlsx"

MASTER_SHEET_NAME = "Sheet1"
TERMINAL_SHEET_NAME = "Terminal"


# =========================================================
# 2. Streamlit page setup
# =========================================================

st.set_page_config(
    page_title="Pickup / Delivery ZIP Heatmap",
    layout="wide"
)

st.title("Pickup / Delivery ZIP Heatmap")


# =========================================================
# Check data file
# =========================================================

if not DATA_FILE.exists():
    st.error(
        f"Cannot find data file: {DATA_FILE}. "
        "Please make sure your Excel file is saved as data/HeatmapData.xlsx."
    )
    st.stop()

# =========================================================
# 3. Helper functions
# =========================================================

def clean_null_text(value):
    """
    Convert Excel NULL / blank values into empty string.
    """
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.upper() in ["NULL", "NAN", "NONE"]:
        return ""

    return value


def clean_zip(value):
    """
    Clean shipment ZIP code.
    Mainly for US 5-digit ZIPs.
    """
    if pd.isna(value):
        return None

    value = str(value).strip().upper()
    value = value.replace("\u00A0", "")

    if value in ["", "NULL", "NAN", "NONE"]:
        return None

    if re.match(r"^\d+\.0$", value):
        value = value.split(".")[0]

    digits = re.sub(r"\D", "", value)

    if len(digits) == 0:
        return None

    if len(digits) < 5:
        return digits.zfill(5)

    return digits[:5]


def clean_terminal(value):
    """
    Clean terminal ID.

    Examples:
    ' AT ' -> 'AT'
    'AT.0' -> 'AT'
    0 / NULL / blank -> None
    """
    if pd.isna(value):
        return None

    value = str(value).strip().upper()
    value = value.replace("\u00A0", "")

    if value.endswith(".0"):
        value = value[:-2]

    if value in ["", "NULL", "NAN", "NONE", "0"]:
        return None

    return value


def clean_terminal_postal(value):
    """
    Clean terminal ZIP / postal code.

    Supports:
    - US ZIP: 30288
    - Canadian postal code: T1X0R2
    """
    if pd.isna(value):
        return None

    value = str(value).strip().upper()
    value = value.replace("\u00A0", "")

    if value in ["", "NULL", "NAN", "NONE"]:
        return None

    value = re.sub(r"[\s\-]", "", value)

    if re.match(r"^\d+\.0$", value):
        value = value.split(".")[0]

    # US ZIP
    if re.match(r"^\d{5}", value):
        return value[:5]

    # Canadian postal code
    if re.match(r"^[A-Z]\d[A-Z]\d[A-Z]\d$", value):
        return value

    return value


def get_terminal_country_from_state(state_value):
    """
    Decide whether terminal postal code is US or Canada based on State column.
    """
    state_value = clean_null_text(state_value).upper()

    canadian_provinces = {
        "AB", "BC", "MB", "NB", "NL", "NS",
        "NT", "NU", "ON", "PE", "QC", "SK", "YT"
    }

    if state_value in canadian_provinces:
        return "CA"

    return "US"


def get_postal_query_code(postal_code, country):
    """
    pgeocode uses:
    - US ZIP: 5-digit ZIP
    - Canada: first 3 characters/FSA works better
    """
    if postal_code is None:
        return None

    postal_code = str(postal_code).replace(" ", "").upper()

    if country == "CA":
        return postal_code[:3]

    return postal_code[:5]


def standardize_shipment_columns(df):
    """
    Standardize Sheet1 / master shipment data.

    Expected columns:
    ProNumber | ShipDate | Zip | TotalWeight | Scac | Type | Terminal
    """

    column_map = {}

    for col in df.columns:
        clean_col = col.strip().lower().replace(" ", "").replace("_", "")

        if clean_col == "pronumber":
            column_map[col] = "ProNumber"
        elif clean_col == "shipdate":
            column_map[col] = "ShipDate"
        elif clean_col in ["zip", "zipcode", "postalcode"]:
            column_map[col] = "Zip"
        elif clean_col == "totalweight":
            column_map[col] = "TotalWeight"
        elif clean_col == "scac":
            column_map[col] = "Scac"
        elif clean_col == "type":
            column_map[col] = "Type"
        elif clean_col in [
            "terminal",
            "terminalid",
            "terminalcode",
            "terminalnumber",
            "term",
            "originterminal"
        ]:
            column_map[col] = "Terminal"

    df = df.rename(columns=column_map)

    required_cols = [
        "ProNumber",
        "ShipDate",
        "Zip",
        "TotalWeight",
        "Scac",
        "Type"
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        st.error(f"Missing required columns in Sheet1: {missing_cols}")
        st.stop()

    if "Terminal" not in df.columns:
        st.warning(
            "No Terminal column found in Sheet1. "
            "The app will use UNKNOWN for terminal."
        )
        df["Terminal"] = "UNKNOWN"

    return df


def standardize_terminal_columns(df):
    """
    Standardize Terminal tab.

    Your terminal tab structure:
    TerminalId | Address | City | State | Zip
    """

    column_map = {}

    for col in df.columns:
        clean_col = col.strip().lower().replace(" ", "").replace("_", "")

        if clean_col in ["terminal", "terminalid", "terminalcode", "term"]:
            column_map[col] = "Terminal"
        elif clean_col == "address":
            column_map[col] = "TerminalAddress"
        elif clean_col == "city":
            column_map[col] = "TerminalCity"
        elif clean_col == "state":
            column_map[col] = "TerminalState"
        elif clean_col in ["zip", "zipcode", "postalcode"]:
            column_map[col] = "TerminalZip"
        elif clean_col == "latitude":
            column_map[col] = "TerminalLatitude"
        elif clean_col == "longitude":
            column_map[col] = "TerminalLongitude"

    df = df.rename(columns=column_map)

    if "Terminal" not in df.columns:
        st.error("Terminal tab must have a TerminalId column.")
        st.stop()

    if "TerminalAddress" not in df.columns:
        df["TerminalAddress"] = ""

    if "TerminalCity" not in df.columns:
        df["TerminalCity"] = ""

    if "TerminalState" not in df.columns:
        df["TerminalState"] = ""

    if "TerminalZip" not in df.columns:
        df["TerminalZip"] = None

    if "TerminalLatitude" not in df.columns:
        df["TerminalLatitude"] = None

    if "TerminalLongitude" not in df.columns:
        df["TerminalLongitude"] = None

    df["Terminal"] = df["Terminal"].apply(clean_terminal)
    df["TerminalAddress"] = df["TerminalAddress"].apply(clean_null_text)
    df["TerminalCity"] = df["TerminalCity"].apply(clean_null_text)
    df["TerminalState"] = df["TerminalState"].apply(clean_null_text).str.upper()
    df["TerminalZip"] = df["TerminalZip"].apply(clean_terminal_postal)

    df["TerminalLatitude"] = pd.to_numeric(df["TerminalLatitude"], errors="coerce")
    df["TerminalLongitude"] = pd.to_numeric(df["TerminalLongitude"], errors="coerce")

    # Remove bad rows such as TerminalId = 0 / NULL
    df = df[df["Terminal"].notna()].copy()

    df["TerminalCountry"] = df["TerminalState"].apply(get_terminal_country_from_state)

    df["PostalQuery"] = df.apply(
        lambda row: get_postal_query_code(row["TerminalZip"], row["TerminalCountry"]),
        axis=1
    )

    return df


@st.cache_data
def load_excel_file(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name)


@st.cache_data
def load_zip_coordinates(zip_list):
    """
    Get latitude and longitude for shipment ZIPs.
    Shipment ZIPs are treated as US ZIPs.
    """

    nomi = pgeocode.Nominatim("us")
    geo = nomi.query_postal_code(zip_list)

    if isinstance(geo, pd.Series):
        geo = geo.to_frame().T

    geo = geo[
        [
            "postal_code",
            "place_name",
            "state_code",
            "county_name",
            "latitude",
            "longitude"
        ]
    ].copy()

    geo = geo.rename(
        columns={
            "postal_code": "Zip",
            "place_name": "City",
            "state_code": "State",
            "county_name": "County",
            "latitude": "Latitude",
            "longitude": "Longitude"
        }
    )

    geo["Zip"] = geo["Zip"].astype(str).str.zfill(5)

    return geo


def prepare_terminal_marker_data(terminal_df, selected_terminals):
    """
    Prepare terminal marker locations.

    If Latitude / Longitude exist in the Terminal tab, use them.
    Otherwise:
    - US terminals use 5-digit ZIP centroid
    - Canadian terminals use first 3-character postal area
    """

    if terminal_df is None or len(terminal_df) == 0:
        return pd.DataFrame()

    terminal_filtered = terminal_df[
        terminal_df["Terminal"].isin(selected_terminals)
    ].copy()

    if len(terminal_filtered) == 0:
        return pd.DataFrame()

    missing_lat_lon = (
        terminal_filtered["TerminalLatitude"].isna() |
        terminal_filtered["TerminalLongitude"].isna()
    )

    need_geo = terminal_filtered[missing_lat_lon].copy()

    geo_results = []

    if len(need_geo) > 0:

        postal_pairs = (
            need_geo[["TerminalCountry", "PostalQuery"]]
            .dropna()
            .drop_duplicates()
        )

        for country in sorted(postal_pairs["TerminalCountry"].unique()):

            country_pairs = postal_pairs[
                postal_pairs["TerminalCountry"] == country
            ].copy()

            query_codes = country_pairs["PostalQuery"].dropna().unique().tolist()

            if len(query_codes) == 0:
                continue

            try:
                nomi = pgeocode.Nominatim(country.lower())
                geo = nomi.query_postal_code(query_codes)

                if isinstance(geo, pd.Series):
                    geo = geo.to_frame().T

                geo = geo[
                    [
                        "postal_code",
                        "latitude",
                        "longitude"
                    ]
                ].copy()

                geo["TerminalCountry"] = country

                geo["PostalQuery"] = geo["postal_code"].astype(str).str.upper()
                geo["PostalQuery"] = geo["PostalQuery"].str.replace(" ", "", regex=False)

                if country == "CA":
                    geo["PostalQuery"] = geo["PostalQuery"].str[:3]
                else:
                    geo["PostalQuery"] = geo["PostalQuery"].str[:5]

                geo = geo.rename(
                    columns={
                        "latitude": "ZipLatitude",
                        "longitude": "ZipLongitude"
                    }
                )

                geo_results.append(
                    geo[
                        [
                            "TerminalCountry",
                            "PostalQuery",
                            "ZipLatitude",
                            "ZipLongitude"
                        ]
                    ]
                )

            except Exception as e:
                st.warning(f"Could not geocode terminal postal codes for {country}: {e}")

    if len(geo_results) > 0:

        geo_lookup = pd.concat(geo_results, ignore_index=True)

        terminal_filtered = terminal_filtered.merge(
            geo_lookup,
            on=["TerminalCountry", "PostalQuery"],
            how="left"
        )

        terminal_filtered["TerminalLatitude"] = terminal_filtered["TerminalLatitude"].fillna(
            terminal_filtered["ZipLatitude"]
        )

        terminal_filtered["TerminalLongitude"] = terminal_filtered["TerminalLongitude"].fillna(
            terminal_filtered["ZipLongitude"]
        )

    terminal_filtered = terminal_filtered.dropna(
        subset=["TerminalLatitude", "TerminalLongitude"]
    ).copy()

    return terminal_filtered


def add_terminal_markers(map_object, terminal_marker_data):
    """
    Add highlighted terminal markers to the map.
    """

    for _, row in terminal_marker_data.iterrows():

        popup_text = f"""
        <b>Selected Terminal:</b> {row['Terminal']}<br>
        <b>Address:</b> {row.get('TerminalAddress', '')}<br>
        <b>City:</b> {row.get('TerminalCity', '')}<br>
        <b>State:</b> {row.get('TerminalState', '')}<br>
        <b>ZIP:</b> {row.get('TerminalZip', '')}
        """

        folium.CircleMarker(
            location=[row["TerminalLatitude"], row["TerminalLongitude"]],
            radius=17,
            color="orange",
            fill=True,
            fill_color="yellow",
            fill_opacity=0.45,
            weight=4,
            popup=folium.Popup(popup_text, max_width=300),
            tooltip=f"Selected Terminal: {row['Terminal']}"
        ).add_to(map_object)

        folium.Marker(
            location=[row["TerminalLatitude"], row["TerminalLongitude"]],
            popup=folium.Popup(popup_text, max_width=300),
            tooltip=f"Selected Terminal: {row['Terminal']}",
            icon=folium.Icon(color="red", icon="home")
        ).add_to(map_object)


def add_zip_circle_marker(
    map_object,
    row,
    color,
    scac_name,
    metric_column,
    max_value,
    lon_offset=0
):
    """
    Add colored ZIP circle marker.
    Used in Compare 2 SCAC mode.
    """

    marker_radius = 6

    if max_value > 0:
        marker_radius = 6 + 18 * row[metric_column] / max_value

    popup_text = f"""
    <b>ZIP:</b> {row['Zip']}<br>
    <b>City:</b> {row.get('City', '')}<br>
    <b>State:</b> {row.get('State', '')}<br>
    <b>County:</b> {row.get('County', '')}<br>
    <b>SCAC:</b> {scac_name}<br>
    <b>Terminal:</b> {row.get('TerminalList', '')}<br>
    <b>Type:</b> {row.get('TypeList', '')}<br>
    <b>Shipment Count:</b> {int(row.get('ShipmentCount', 0))}<br>
    <b>Total Weight:</b> {row.get('TotalWeight', 0):,.0f}
    """

    folium.CircleMarker(
        location=[row["Latitude"], row["Longitude"] + lon_offset],
        radius=marker_radius,
        popup=folium.Popup(popup_text, max_width=300),
        tooltip=f"{scac_name} | ZIP {row['Zip']} | {metric_column}: {row[metric_column]:,.0f}",
        color=color,
        fill=True,
        fill_color=color,
        fill_opacity=0.65,
        weight=2
    ).add_to(map_object)


# =========================================================
# 4. Find Excel files
# =========================================================

excel_files = sorted(
    [
        file for file in BASE_DIR.glob("*.xlsx")
        if not file.name.startswith("~$")
    ]
)

if not excel_files:
    st.error(f"No Excel file found in this folder: {BASE_DIR}")
    st.stop()

selected_file = st.sidebar.selectbox(
    "Select Excel File",
    options=excel_files,
    format_func=lambda x: x.name
)


# =========================================================
# 5. Load fixed sheets: Sheet1 and Terminal
# =========================================================

try:
    master_raw = load_excel_file(str(selected_file), MASTER_SHEET_NAME)
except Exception as e:
    st.error(f"Could not load sheet '{MASTER_SHEET_NAME}'. Error: {e}")
    st.stop()

try:
    terminal_raw = load_excel_file(str(selected_file), TERMINAL_SHEET_NAME)
except Exception as e:
    st.error(f"Could not load sheet '{TERMINAL_SHEET_NAME}'. Error: {e}")
    st.stop()


# =========================================================
# 6. Clean master data
# =========================================================

df = standardize_shipment_columns(master_raw)

df["ProNumber"] = df["ProNumber"].astype(str).str.strip()
df["ShipDate"] = pd.to_datetime(df["ShipDate"], errors="coerce")
df["Zip"] = df["Zip"].apply(clean_zip)
df["TotalWeight"] = pd.to_numeric(df["TotalWeight"], errors="coerce").fillna(0)
df["Scac"] = df["Scac"].astype(str).str.strip().str.upper()
df["Type"] = df["Type"].astype(str).str.strip().str.upper()
df["Terminal"] = df["Terminal"].apply(clean_terminal)
df["Terminal"] = df["Terminal"].fillna("UNKNOWN")

df = df[df["Type"].isin(["P", "D"])].copy()
df = df[df["Zip"].notna()].copy()


# =========================================================
# 7. Clean terminal tab
# =========================================================

terminal_df = standardize_terminal_columns(terminal_raw)


# =========================================================
# 8. Match terminal tab back to master data
# =========================================================

terminal_lookup = terminal_df[
    [
        "Terminal",
        "TerminalAddress",
        "TerminalCity",
        "TerminalState",
        "TerminalZip",
        "TerminalCountry",
        "TerminalLatitude",
        "TerminalLongitude"
    ]
].drop_duplicates(subset=["Terminal"]).copy()

df = df.merge(
    terminal_lookup,
    on="Terminal",
    how="left",
    indicator=True
)

unmatched_terminals = sorted(
    df.loc[
        (df["_merge"] == "left_only") &
        (df["Terminal"].notna()) &
        (df["Terminal"] != "UNKNOWN"),
        "Terminal"
    ].unique()
)

if len(unmatched_terminals) > 0:
    st.warning(
        "These terminal IDs exist in Sheet1 but were not found in the Terminal tab: "
        f"{unmatched_terminals}"
    )

df = df.drop(columns=["_merge"])


# =========================================================
# 9. Sidebar filters
# =========================================================

st.sidebar.header("Filters")

map_mode = st.sidebar.radio(
    "Map Mode",
    options=[
        "Normal Heatmap",
        "Compare 2 SCAC"
    ]
)

type_options = sorted(df["Type"].dropna().unique())

selected_types = st.sidebar.multiselect(
    "Pickup / Delivery Type",
    options=type_options,
    default=type_options
)

terminal_options = sorted(df["Terminal"].dropna().unique())

selected_terminals = st.sidebar.multiselect(
    "Terminal",
    options=terminal_options,
    default=terminal_options
)

if len(selected_terminals) == 0:
    st.warning("Please select at least one terminal.")
    st.stop()

metric_choice = st.sidebar.radio(
    "Map Weight",
    options=["Shipment Count", "Total Weight"]
)

map_weight_column = "ShipmentCount" if metric_choice == "Shipment Count" else "TotalWeight"

show_terminal_markers = st.sidebar.checkbox(
    "Show Selected Terminal Markers",
    value=True
)

valid_dates = df["ShipDate"].dropna()

if len(valid_dates) > 0:
    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()

    selected_date_range = st.sidebar.date_input(
        "Ship Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    selected_date_range = None

zip_filter_text = st.sidebar.text_input(
    "Optional ZIP Filter",
    placeholder="Example: 95820, 95242, 95035"
)


# =========================================================
# 10. Base filters
# =========================================================

filtered_base = df[
    (df["Type"].isin(selected_types)) &
    (df["Terminal"].isin(selected_terminals))
].copy()

if selected_date_range and len(selected_date_range) == 2:
    start_date = pd.Timestamp(selected_date_range[0])
    end_date = pd.Timestamp(selected_date_range[1]) + pd.Timedelta(days=1)

    filtered_base = filtered_base[
        (filtered_base["ShipDate"] >= start_date) &
        (filtered_base["ShipDate"] < end_date)
    ].copy()

if zip_filter_text.strip():
    selected_zips = [
        clean_zip(x)
        for x in zip_filter_text.split(",")
        if clean_zip(x) is not None
    ]

    filtered_base = filtered_base[
        filtered_base["Zip"].isin(selected_zips)
    ].copy()

terminal_marker_data = prepare_terminal_marker_data(
    terminal_df=terminal_df,
    selected_terminals=selected_terminals
)


# =========================================================
# 11A. Normal Heatmap Mode
# =========================================================

if map_mode == "Normal Heatmap":

    scac_options = sorted(filtered_base["Scac"].dropna().unique())

    selected_scacs = st.sidebar.multiselect(
        "SCAC",
        options=scac_options,
        default=scac_options
    )

    filtered = filtered_base[
        filtered_base["Scac"].isin(selected_scacs)
    ].copy()

    if len(filtered) == 0:
        st.warning("No data found after applying filters.")
        st.stop()

    detail_summary = (
        filtered
        .groupby(["Terminal", "Zip", "Type", "Scac"], as_index=False)
        .agg(
            ShipmentCount=("ProNumber", "nunique"),
            TotalWeight=("TotalWeight", "sum")
        )
    )

    map_summary = (
        filtered
        .groupby("Zip", as_index=False)
        .agg(
            ShipmentCount=("ProNumber", "nunique"),
            TotalWeight=("TotalWeight", "sum"),
            ScacList=("Scac", lambda x: ", ".join(sorted(x.dropna().unique()))),
            TypeList=("Type", lambda x: ", ".join(sorted(x.dropna().unique()))),
            TerminalList=("Terminal", lambda x: ", ".join(sorted(x.dropna().unique())))
        )
    )

    zip_list = sorted(map_summary["Zip"].dropna().unique().tolist())
    zip_geo = load_zip_coordinates(zip_list)

    map_data = map_summary.merge(
        zip_geo,
        on="Zip",
        how="left"
    )

    missing_geo = map_data[
        map_data["Latitude"].isna() |
        map_data["Longitude"].isna()
    ]

    map_data = map_data.dropna(subset=["Latitude", "Longitude"]).copy()

    if len(missing_geo) > 0:
        st.warning(f"{len(missing_geo)} ZIP code(s) could not be mapped and were excluded.")

    st.subheader("Summary")

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Total Shipments", int(filtered["ProNumber"].nunique()))
    col2.metric("Total Weight", f"{filtered['TotalWeight'].sum():,.0f}")
    col3.metric("Unique ZIPs", filtered["Zip"].nunique())
    col4.metric("Unique SCACs", filtered["Scac"].nunique())
    col5.metric("Unique Terminals", filtered["Terminal"].nunique())

    st.subheader("ZIP Heatmap")

    center_lat = map_data["Latitude"].mean()
    center_lon = map_data["Longitude"].mean()

    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=7,
        tiles="CartoDB positron"
    )

    heat_data = map_data[
        ["Latitude", "Longitude", map_weight_column]
    ].values.tolist()

    HeatMap(
        heat_data,
        radius=22,
        blur=16,
        min_opacity=0.35,
        name="Heatmap"
    ).add_to(m)

    max_value = map_data[map_weight_column].max()

    for _, row in map_data.iterrows():

        marker_radius = 5

        if max_value > 0:
            marker_radius = 5 + 15 * row[map_weight_column] / max_value

        popup_text = f"""
        <b>ZIP:</b> {row['Zip']}<br>
        <b>City:</b> {row.get('City', '')}<br>
        <b>State:</b> {row.get('State', '')}<br>
        <b>County:</b> {row.get('County', '')}<br>
        <b>Terminal:</b> {row['TerminalList']}<br>
        <b>Type:</b> {row['TypeList']}<br>
        <b>SCAC:</b> {row['ScacList']}<br>
        <b>Shipment Count:</b> {int(row['ShipmentCount'])}<br>
        <b>Total Weight:</b> {row['TotalWeight']:,.0f}
        """

        folium.CircleMarker(
            location=[row["Latitude"], row["Longitude"]],
            radius=marker_radius,
            popup=folium.Popup(popup_text, max_width=300),
            fill=True,
            fill_opacity=0.65,
            weight=1
        ).add_to(m)

    if show_terminal_markers and len(terminal_marker_data) > 0:
        add_terminal_markers(m, terminal_marker_data)

    folium.LayerControl().add_to(m)

    st_folium(m, width=1300, height=700)

    st.subheader("ZIP Summary for Map")

    map_display = map_data[
        [
            "Zip",
            "City",
            "State",
            "County",
            "TerminalList",
            "TypeList",
            "ScacList",
            "ShipmentCount",
            "TotalWeight",
            "Latitude",
            "Longitude"
        ]
    ].sort_values(map_weight_column, ascending=False)

    st.dataframe(
        map_display,
        use_container_width=True
    )

    st.subheader("Detail Summary by Terminal / ZIP / Type / SCAC")

    detail_display = detail_summary.sort_values(
        ["Terminal", "ShipmentCount", "TotalWeight"],
        ascending=[True, False, False]
    )

    st.dataframe(
        detail_display,
        use_container_width=True
    )

    st.download_button(
        label="Download ZIP Summary CSV",
        data=map_display.to_csv(index=False),
        file_name="zip_heatmap_summary.csv",
        mime="text/csv"
    )

    st.download_button(
        label="Download Detail Summary CSV",
        data=detail_display.to_csv(index=False),
        file_name="terminal_zip_type_scac_detail_summary.csv",
        mime="text/csv"
    )


# =========================================================
# 11B. Compare 2 SCAC Mode
# =========================================================

else:

    scac_options = sorted(filtered_base["Scac"].dropna().unique())

    if len(scac_options) < 2:
        st.warning("You need at least two SCACs in the filtered data to use Compare 2 SCAC mode.")
        st.stop()

    scac_a = st.sidebar.selectbox(
        "SCAC A",
        options=scac_options,
        index=0
    )

    scac_b_options = [x for x in scac_options if x != scac_a]

    scac_b = st.sidebar.selectbox(
        "SCAC B",
        options=scac_b_options,
        index=0
    )

    compare_filtered = filtered_base[
        filtered_base["Scac"].isin([scac_a, scac_b])
    ].copy()

    if len(compare_filtered) == 0:
        st.warning("No data found for selected SCACs.")
        st.stop()

    compare_summary = (
        compare_filtered
        .groupby(["Zip", "Scac"], as_index=False)
        .agg(
            ShipmentCount=("ProNumber", "nunique"),
            TotalWeight=("TotalWeight", "sum"),
            TypeList=("Type", lambda x: ", ".join(sorted(x.dropna().unique()))),
            TerminalList=("Terminal", lambda x: ", ".join(sorted(x.dropna().unique())))
        )
    )

    zip_list = sorted(compare_summary["Zip"].dropna().unique().tolist())
    zip_geo = load_zip_coordinates(zip_list)

    compare_map_data = compare_summary.merge(
        zip_geo,
        on="Zip",
        how="left"
    )

    compare_map_data = compare_map_data.dropna(
        subset=["Latitude", "Longitude"]
    ).copy()

    if len(compare_map_data) == 0:
        st.warning("No ZIPs with valid latitude and longitude.")
        st.stop()

    count_pivot = compare_summary.pivot_table(
        index="Zip",
        columns="Scac",
        values="ShipmentCount",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    weight_pivot = compare_summary.pivot_table(
        index="Zip",
        columns="Scac",
        values="TotalWeight",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    terminal_list_by_zip = (
        compare_summary
        .groupby("Zip", as_index=False)
        .agg(
            TerminalList=("TerminalList", lambda x: ", ".join(sorted(set(", ".join(x).split(", "))))),
            TypeList=("TypeList", lambda x: ", ".join(sorted(set(", ".join(x).split(", ")))))
        )
    )

    for scac in [scac_a, scac_b]:
        if scac not in count_pivot.columns:
            count_pivot[scac] = 0

        if scac not in weight_pivot.columns:
            weight_pivot[scac] = 0

    comparison_table = count_pivot[["Zip", scac_a, scac_b]].copy()

    comparison_table = comparison_table.rename(
        columns={
            scac_a: f"{scac_a}_ShipmentCount",
            scac_b: f"{scac_b}_ShipmentCount"
        }
    )

    comparison_table = comparison_table.merge(
        weight_pivot[["Zip", scac_a, scac_b]].rename(
            columns={
                scac_a: f"{scac_a}_TotalWeight",
                scac_b: f"{scac_b}_TotalWeight"
            }
        ),
        on="Zip",
        how="left"
    )

    comparison_table[f"{scac_a}_Minus_{scac_b}_ShipmentCount"] = (
        comparison_table[f"{scac_a}_ShipmentCount"] -
        comparison_table[f"{scac_b}_ShipmentCount"]
    )

    comparison_table[f"{scac_a}_Minus_{scac_b}_TotalWeight"] = (
        comparison_table[f"{scac_a}_TotalWeight"] -
        comparison_table[f"{scac_b}_TotalWeight"]
    )

    def coverage_status(row):
        a_count = row[f"{scac_a}_ShipmentCount"]
        b_count = row[f"{scac_b}_ShipmentCount"]

        if a_count > 0 and b_count > 0:
            return "Both"
        elif a_count > 0 and b_count == 0:
            return f"Only {scac_a}"
        elif a_count == 0 and b_count > 0:
            return f"Only {scac_b}"
        else:
            return "None"

    comparison_table["CoverageStatus"] = comparison_table.apply(
        coverage_status,
        axis=1
    )

    comparison_table = comparison_table.merge(
        terminal_list_by_zip,
        on="Zip",
        how="left"
    )

    comparison_table = comparison_table.merge(
        zip_geo,
        on="Zip",
        how="left"
    )

    st.subheader("SCAC Comparison Summary")

    a_total_count = comparison_table[f"{scac_a}_ShipmentCount"].sum()
    b_total_count = comparison_table[f"{scac_b}_ShipmentCount"].sum()

    both_zip_count = (comparison_table["CoverageStatus"] == "Both").sum()
    only_a_zip_count = (comparison_table["CoverageStatus"] == f"Only {scac_a}").sum()
    only_b_zip_count = (comparison_table["CoverageStatus"] == f"Only {scac_b}").sum()

    col1, col2, col3, col4, col5, col6 = st.columns(6)

    col1.metric(f"{scac_a} Shipments", int(a_total_count))
    col2.metric(f"{scac_b} Shipments", int(b_total_count))
    col3.metric("ZIPs Covered by Both", int(both_zip_count))
    col4.metric(f"Only {scac_a} ZIPs", int(only_a_zip_count))
    col5.metric(f"Only {scac_b} ZIPs", int(only_b_zip_count))
    col6.metric("Selected Terminals", len(selected_terminals))

    st.markdown(
        f"""
        **Map Color Rule:**  
        <span style="color:blue;"><b>{scac_a}</b></span> = Blue  
        &nbsp;&nbsp;&nbsp;  
        <span style="color:red;"><b>{scac_b}</b></span> = Red  
        """,
        unsafe_allow_html=True
    )

    st.subheader("Compare 2 SCAC Map")

    center_lat = compare_map_data["Latitude"].mean()
    center_lon = compare_map_data["Longitude"].mean()

    m = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=7,
        tiles="CartoDB positron"
    )

    scac_a_data = compare_map_data[
        compare_map_data["Scac"] == scac_a
    ].copy()

    scac_b_data = compare_map_data[
        compare_map_data["Scac"] == scac_b
    ].copy()

    max_value = compare_map_data[map_weight_column].max()

    if len(scac_a_data) > 0:
        heat_a = scac_a_data[
            ["Latitude", "Longitude", map_weight_column]
        ].values.tolist()

        HeatMap(
            heat_a,
            radius=22,
            blur=16,
            min_opacity=0.35,
            gradient={
                0.3: "lightblue",
                0.6: "blue",
                1.0: "darkblue"
            },
            name=f"{scac_a} Heatmap"
        ).add_to(m)

    if len(scac_b_data) > 0:
        heat_b = scac_b_data[
            ["Latitude", "Longitude", map_weight_column]
        ].values.tolist()

        HeatMap(
            heat_b,
            radius=22,
            blur=16,
            min_opacity=0.35,
            gradient={
                0.3: "pink",
                0.6: "red",
                1.0: "darkred"
            },
            name=f"{scac_b} Heatmap"
        ).add_to(m)

    for _, row in scac_a_data.iterrows():
        add_zip_circle_marker(
            map_object=m,
            row=row,
            color="blue",
            scac_name=scac_a,
            metric_column=map_weight_column,
            max_value=max_value,
            lon_offset=-0.01
        )

    for _, row in scac_b_data.iterrows():
        add_zip_circle_marker(
            map_object=m,
            row=row,
            color="red",
            scac_name=scac_b,
            metric_column=map_weight_column,
            max_value=max_value,
            lon_offset=0.01
        )

    if show_terminal_markers and len(terminal_marker_data) > 0:
        add_terminal_markers(m, terminal_marker_data)

    folium.LayerControl().add_to(m)

    st_folium(m, width=1300, height=700)

    st.subheader("SCAC ZIP Comparison Table")

    comparison_display_cols = [
        "Zip",
        "City",
        "State",
        "County",
        "TerminalList",
        "TypeList",
        f"{scac_a}_ShipmentCount",
        f"{scac_b}_ShipmentCount",
        f"{scac_a}_Minus_{scac_b}_ShipmentCount",
        f"{scac_a}_TotalWeight",
        f"{scac_b}_TotalWeight",
        f"{scac_a}_Minus_{scac_b}_TotalWeight",
        "CoverageStatus",
        "Latitude",
        "Longitude"
    ]

    comparison_display = comparison_table[
        comparison_display_cols
    ].sort_values(
        by=[
            "CoverageStatus",
            f"{scac_a}_ShipmentCount",
            f"{scac_b}_ShipmentCount"
        ],
        ascending=[True, False, False]
    )

    st.dataframe(
        comparison_display,
        use_container_width=True
    )

    st.download_button(
        label="Download SCAC Comparison CSV",
        data=comparison_display.to_csv(index=False),
        file_name=f"scac_comparison_{scac_a}_vs_{scac_b}.csv",
        mime="text/csv"
    )