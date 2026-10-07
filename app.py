import io
import pandas as pd
import pyreadstat
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="SME-ENT Tables Generator",
    page_icon="📊",
    layout="wide",
)

st.title("SME-ENT Tables: R10MIL_GROWTH")
st.write(
    "Upload your SPSS `.sav` data file below. Filters for Type = 1 (Growth)"
    " & 2 (R10Mil), and calculates metrics across Total, Type, and Segment"
    " banners."
)

# File uploader widget
uploaded_file = st.file_uploader(
    "Upload .SAV File", type=["sav"], help="Upload your SPSS dataset here."
)

# Execution button
if uploaded_file is not None:
  if st.button("Run Extraction"):
    with st.spinner("Processing dataset... Please wait."):
      # Save uploaded file temporarily so pyreadstat can read it
      with open("temp.sav", "wb") as f:
        f.write(uploaded_file.getbuffer())

      # Read the .sav file and metadata using pyreadstat
      df, meta = pyreadstat.read_sav("temp.sav")

      # Normalize column names for case-insensitive matching if needed
      # Find exact column names regardless of casing
      col_map = {c.lower(): c for c in df.columns}

      type_col = col_map.get("type", None)
      seg2_col = col_map.get("seg2", None)

      # --- FILTER LOGIC FOR R10MIL_GROWTH ---
      if type_col:
        df_tab = df[df[type_col].isin([1, 2])]
      else:
        st.error(
            "Error: 'Type' column not found in the uploaded SPSS file."
            " Please check your variable names."
        )
        df_tab = pd.DataFrame()

      if not df_tab.empty:

        # Helper: Determine segment codes dynamically from metadata if possible
        ent_val, plat_val = 1, 2  # defaults
        if (
            meta
            and meta.variable_value_labels
            and seg2_col
            and seg2_col in meta.variable_value_labels
        ):
          seg2_labels = meta.variable_value_labels[seg2_col]
          ent_code = [
              k for k, v in seg2_labels.items() if "enterprise" in str(v).lower()
          ]
          plat_code = [
              k for k, v in seg2_labels.items() if "platinum" in str(v).lower()
          ]
          if ent_code:
            ent_val = ent_code[0]
          if plat_code:
            plat_val = plat_code[0]


        # Function to calculate NPS Score and Valid N
        def calculate_nps(data, variable_name):
          if variable_name not in data.columns or len(data) == 0:
            return 0.0, 0
          valid_data = data[data[variable_name].isin([1, 2, 3])]
          valid_n = len(valid_data)
          if valid_n == 0:
            return 0.0, 0
          promoters = (valid_data[variable_name] == 3).sum()
          detractors = (valid_data[variable_name] == 1).sum()
          nps_score = ((promoters - detractors) / valid_n) * 100
          return round(nps_score, 2), valid_n


        # Function to calculate Rating Mean (excluding code 11 for "Don't know")
        def calculate_rating_mean(data, variable_name):
          if variable_name not in data.columns or len(data) == 0:
            return 0.0
          valid_data = data[
              (data[variable_name].notnull()) & (data[variable_name] != 11)
          ]
          if len(valid_data) == 0:
            return 0.0
          return round(valid_data[variable_name].mean(), 2)


        # Define sub-samples for the columns safely using detected column names
        seg2_filter_col = seg2_col if seg2_col else "seg2"

        subsets = {
            ("Total", "Total"): df_tab,
            ("Type", "Growth"): (
                df_tab[df_tab[type_col] == 1] if type_col else pd.DataFrame()
            ),
            ("Type", "R10Mil"): (
                df_tab[df_tab[type_col] == 2] if type_col else pd.DataFrame()
            ),
            ("Segment", "ENTERPRISE"): (
                df_tab[df_tab[seg2_filter_col] == ent_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            ),
            ("Segment", "PLATINUM"): (
                df_tab[df_tab[seg2_filter_col] == plat_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            ),
        }

        # Define metrics to extract (NPS, Q10, Q11, Q12)
        metrics_config = [
            ("FNB_NPS_SCORE", "FNB_NPS1", True),
            ("BM_NPS_SCORE", "BM_NPS1", True),
            ("Q10.1. OVERALL BRANCH EXPERIENCE?", "Q10_2", False),
            ("Q10.2. OVERALL CONTACT CENTRE EXPERIENCE?", "Q10_3", False),
            ("Q10.3. OVERALL ONLINE BANKING EXPERIENCE?", "Q10_4", False),
            (
                "Q10.4. OVERALL FNB BUSINESS BANKING APP EXPERIENCE?",
                "Q10_5",
                False,
            ),
            ("Q11.1 Lending products", "Q11_1_1", False),
            ("Q11.2 Transactional products", "Q11_1_2", False),
            ("Q11.3 Insurance products", "Q11_1_3", False),
            ("Q11.4 Investment products", "Q11_1_4", False),
            ("Q11.5 Forex Products", "Q11_1_5", False),
            (
                (
                    "Q12. Your overall level of satisfaction with the products"
                    " you received from FNB Business?"
                ),
                "Q12_1",
                False,
            ),
            (
                (
                    "Q12. Your overall level of satisfaction with FNB Business"
                    " over the last 3 months?"
                ),
                "Q12_2",
                False,
            ),
            (
                (
                    "Q12.2 Your overall level of satisfaction with your"
                    " Relationship Manager over the last 3-6 months?"
                ),
                "Q12_3",
                False,
            ),
        ]

        # Build rows data
        table_rows = []
        for label, var_name, is_nps in metrics_config:
          row_data = {"Metric": label}
          for col_key, sub_df in subsets.items():
            banner, sub_col = col_key
            if is_nps:
              mean_val, n_val = calculate_nps(sub_df, var_name)
              row_data[(banner, sub_col, "Mean")] = mean_val
              row_data[(banner, sub_col, "Valid N")] = n_val
            else:
              mean_val = calculate_rating_mean(sub_df, var_name)
              row_data[(banner, sub_col, "Mean")] = mean_val
              row_data[(banner, sub_col, "Valid N")] = (
                  ""  # Blank for rating scales
              )
          table_rows.append(row_data)

        # Construct MultiIndex columns dataframe
        multi_cols = pd.MultiIndex.from_tuples(
            [("Metric", "", "")]
            + [
                (banner, sub_col, stat)
                for banner, sub_col in subsets.keys()
                for stat in ["Mean", "Valid N"]
            ],
            names=["Banner", "Sub-Group", "Stat"],
        )

        # Flatten rows into a standard dataframe structure with MultiIndex columns
        formatted_rows = []
        for r in table_rows:
          flat_row = [r["Metric"]]
          for banner, sub_col in subsets.keys():
            flat_row.append(r[(banner, sub_col, "Mean")])
            flat_row.append(r[(banner, sub_col, "Valid N")])
          formatted_rows.append(flat_row)

        summary_df = pd.DataFrame(formatted_rows, columns=multi_cols)

        st.success(
            "Extraction and banner cross-tabulation completed successfully!"
        )

        # Display dataframe in app
        st.subheader("Results Preview: R10MIL_GROWTH")
        st.dataframe(summary_df, use_container_width=True)

        # Generate Excel file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          summary_df.to_excel(
              writer, sheet_name="R10MIL_GROWTH", index=False, header=True
          )
        excel_data = output.getvalue()

        # Download button
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name="R10MIL_GROWTH_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
