import streamlit as st
import pandas as pd

st.set_page_config(page_title="Color Clarity Automation")

st.title("Color Clarity Automation")

# =========================
# FILE UPLOADERS
# =========================

color_clarity_file = st.file_uploader(
    "Upload Color Clarity File",
    type=["xlsx"]
)

pending_video_file = st.file_uploader(
    "Upload Pending Video File",
    type=["xlsx"]
)

# =========================
# PROCESS FILES
# =========================

if color_clarity_file and pending_video_file:

    # =========================
    # LOAD FILES
    # =========================

    df1 = pd.read_excel(color_clarity_file)

    # Read SECOND SHEET
    df2 = pd.read_excel(pending_video_file, sheet_name=0)

    # =========================
    # STEP 1 - KEEP ONLY IGI
    # =========================

    df1 = df1[
        df1["Lab"].astype(str).str.upper() == "IGI"
    ]

    # =========================
    # STEP 2 - COLOR CLEANING
    # =========================

    allowed_colors = [
        "BLUE",
        "BROWN",
        "YELLOW",
        "GREEN",
        "ORANGE",
        "PINK"
    ]

    def clean_color(color):

        if pd.isna(color):
            return None

        color_upper = str(color).upper()

        for allowed in allowed_colors:

            if allowed in color_upper:
                return allowed

        return None

    df1["Color"] = df1["Color"].apply(clean_color)

    # Remove unmatched colors
    df1 = df1[
        df1["Color"].notna()
    ]

    # =========================
    # STEP 3 - KEEP REQUIRED COLUMNS
    # =========================

    required_columns = [
        "Serial #",
        "Shape",
        "Color",
        "Grade",
        "Cts.",
        "Lab",
        "H&&A",
        "Certificate #"
    ]

    df1 = df1[required_columns]

    # =========================
    # STEP 4 - SHAPE CLEANING
    # =========================

    df1["Shape"] = df1["Shape"].replace({
        "CUSHION MODIFIED": "CUSHION",
        "CUSHION BRILLIANT": "CUSHION"
    })

    # =========================
    # STEP 5 - REMOVE SERIAL PREFIXES
    # =========================

    remove_prefixes = (
        "NY",
        "V",
        "DM",
        "DC"
    )

    df1 = df1[
        ~df1["Serial #"]
        .astype(str)
        .str.upper()
        .str.startswith(remove_prefixes)
    ]

    # =========================
    # STEP 6 - CUSTOMER STATUS LOGIC
    # =========================

    customer_values = [
        "GOODS IN TRANSIT FROM OVERSEAS",
        "GOODS IN OFFICE - PARCEL PAPERS BEING MADE",
        "JCK 2026"
    ]

    mask = (
        df2["Customer"]
        .astype(str)
        .str.upper()
        .isin(customer_values)
    )

    df2.loc[
        mask &
        (
            df2["Status"]
            .astype(str)
            .str.upper() == "ONMEMO"
        ),
        "Status"
    ] = "Inhand"

    # =========================
    # STEP 7 - MATCH SERIAL # WITH LOT #
    # =========================

    status_map = (
        df2
        .set_index("Lot #")["Status"]
        .to_dict()
    )

    serial_index = df1.columns.get_loc("Serial #")

    df1.insert(
        serial_index + 1,
        "Status",
        df1["Serial #"].map(status_map)
    )

    # =========================
    # STEP 8 - SIZE GROUP
    # =========================

    def get_size_group(cts):

        try:

            cts = float(cts)

            if 1.00 <= cts <= 1.49:
                return "1.00-1.49"

            elif 1.50 <= cts <= 1.99:
                return "1.50-1.99"

            elif 2.00 <= cts <= 2.99:
                return "2.00-2.99"

            elif 3.00 <= cts <= 3.99:
                return "3.00-3.99"

            elif 4.00 <= cts <= 4.99:
                return "4.00-4.99"

            elif 5.00 <= cts <= 5.99:
                return "5.00-5.99"

            elif 6.00 <= cts <= 6.99:
                return "6.00-6.99"

            elif 7.00 <= cts <= 7.99:
                return "7.00-7.99"

            elif 8.00 <= cts <= 8.99:
                return "8.00-8.99"

            elif 9.00 <= cts <= 9.99:
                return "9.00-9.99"

            elif 10.00 <= cts <= 10.99:
                return "10.00-10.99"

            else:
                return "Out of Range"

        except:
            return "Out of Range"

    # Add SIZE GROUP column
    cts_index = df1.columns.get_loc("Cts.")

    df1.insert(
        cts_index + 1,
        "SIZE GROUP",
        df1["Cts."].apply(get_size_group)
    )

    # Remove Out of Range
    df1 = df1[
        df1["SIZE GROUP"] != "Out of Range"
    ]

    # =========================
    # DOWNLOAD OUTPUT
    # =========================

    output_file = "Final_Output.xlsx"

    df1.to_excel(
        output_file,
        index=False
    )

    st.success("Automation Completed Successfully!")

    with open(output_file, "rb") as file:

        st.download_button(
            label="Download Final Output",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
