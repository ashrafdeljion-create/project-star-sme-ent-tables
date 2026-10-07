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
    "Upload your SPSS `.sav` data file below. This will filter for Type = 1"
    " (Growth) and Type = 2 (R10Mil) and calculate the initial NPS metrics."
)

# File uploader widget matching your UI design
uploaded_file = st.file_uploader(
    "Upload .SAV File", type=["sav"], help="Upload your SPSS dataset here."
)

# Execution button
if uploaded_file is not None:
  if st.button("Run NPS Extraction"):
    with st.spinner("Processing dataset... Please wait."):
      # Save uploaded file temporarily so pyreadstat can read it
      with open("temp.sav", "wb") as f:
        f.write(uploaded_file.getbuffer())

      # Read the .sav file using pyreadstat
      df, meta = pyreadstat.read_sav("temp.sav")

      # --- FILTER LOGIC FOR R10MIL_GROWTH ---
      # Check if 'Type' column exists in the dataset
      if "Type" in df.columns:
        # Filter for Growth (1) and R10Mil (2)
        df_tab = df[df["Type"].isin([1, 2])]
      else:
        st.error(
            "Error: 'Type' column not found in the uploaded SPSS file."
            " Please check your variable names."
        )
        df_tab = pd.DataFrame()

      if not df_tab.empty:


        # Function to calculate NPS Score and Valid N
        # Code 3 = Promoters, Code 1 = Detractors, Code 2 = Passives
        def calculate_nps(data, variable_name):
          if variable_name not in data.columns:
            return 0.0, 0

          # Filter for valid NPS codes (1, 2, 3)
          valid_data = data[data[variable_name].isin([1, 2, 3])]
          valid_n = len(valid_data)

          if valid_n == 0:
            return 0.0, 0

          promoters = (valid_data[variable_name] == 3).sum()
          detractors = (valid_data[variable_name] == 1).sum()

          # NPS formula: (% Promoters - % Detractors) * 100
          nps_score = ((promoters - detractors) / valid_n) * 100
          return round(nps_score, 2), valid_n


        # Calculate metrics for FNB and BM NPS variables
        fnb_mean, fnb_n = calculate_nps(df_tab, "FNB_NPS1")
        bm_mean, bm_n = calculate_nps(df_tab, "BM_NPS1")

        # Build the output summary table matching your table structure
        summary_df = pd.DataFrame({
            "Metric": ["FNB_NPS_SCORE", "BM_NPS_SCORE"],
            "Total - Mean": [fnb_mean, bm_mean],
            "Total - Valid N": [fnb_n, bm_n],
        })

        st.success(
            "Extraction completed successfully! (Filtered for Type = Growth &"
            " R10Mil)"
        )

        # Display the dataframe in the app
        st.subheader("Results Preview: R10MIL_GROWTH")
        st.dataframe(summary_df, use_container_width=True)

        # Generate Excel file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          summary_df.to_excel(writer, sheet_name="R10MIL_GROWTH", index=False)
        excel_data = output.getvalue()

        # Download button
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name="R10MIL_GROWTH_NPS_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
