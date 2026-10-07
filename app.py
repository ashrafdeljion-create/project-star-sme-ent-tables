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
    " and Type = 2 (R10Mil) and calculates NPS, Q10, Q11, and Q12 Means."
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

      # Read the .sav file using pyreadstat
      df, meta = pyreadstat.read_sav("temp.sav")

      # --- FILTER LOGIC FOR R10MIL_GROWTH ---
      if "Type" in df.columns:
        df_tab = df[df["Type"].isin([1, 2])]
      else:
        st.error(
            "Error: 'Type' column not found in the uploaded SPSS file."
            " Please check your variable names."
        )
        df_tab = pd.DataFrame()

      if not df_tab.empty:


        # Function to calculate NPS Score and Valid N
        def calculate_nps(data, variable_name):
          if variable_name not in data.columns:
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
          if variable_name not in data.columns:
            return 0.0

          valid_data = data[
              (data[variable_name].notnull()) & (data[variable_name] != 11)
          ]
          if len(valid_data) == 0:
            return 0.0

          mean_val = valid_data[variable_name].mean()
          return round(mean_val, 2)


        # 1. Calculate NPS Metrics
        fnb_mean, fnb_n = calculate_nps(df_tab, "FNB_NPS1")
        bm_mean, bm_n = calculate_nps(df_tab, "BM_NPS1")

        # 2. Calculate Q10 Means
        q10_2_mean = calculate_rating_mean(df_tab, "Q10_2")
        q10_3_mean = calculate_rating_mean(df_tab, "Q10_3")
        q10_4_mean = calculate_rating_mean(df_tab, "Q10_4")
        q10_5_mean = calculate_rating_mean(df_tab, "Q10_5")

        # 3. Calculate Q11 Product Means[cite: 6, 7]
        q11_1_mean = calculate_rating_mean(df_tab, "Q11_1_1")
        q11_2_mean = calculate_rating_mean(df_tab, "Q11_1_2")
        q11_3_mean = calculate_rating_mean(df_tab, "Q11_1_3")
        q11_4_mean = calculate_rating_mean(df_tab, "Q11_1_4")
        q11_5_mean = calculate_rating_mean(df_tab, "Q11_1_5")

        # 4. Calculate Q12 Satisfaction Means[cite: 6, 7]
        q12_1_mean = calculate_rating_mean(df_tab, "Q12_1")
        q12_2_mean = calculate_rating_mean(df_tab, "Q12_2")
        q12_3_mean = calculate_rating_mean(df_tab, "Q12_3")

        # Build the output summary table
        summary_data = {
            "Metric": [
                "FNB_NPS_SCORE",
                "BM_NPS_SCORE",
                "Q10.1. OVERALL BRANCH EXPERIENCE?",
                "Q10.2. OVERALL CONTACT CENTRE EXPERIENCE?",
                "Q10.3. OVERALL ONLINE BANKING EXPERIENCE?",
                "Q10.4. OVERALL FNB BUSINESS BANKING APP EXPERIENCE?",
                "Q11.1 Lending products",
                "Q11.2 Transactional products",
                "Q11.3 Insurance products",
                "Q11.4 Investment products",
                "Q11.5 Forex Products",
                (
                    "Q12. Your overall level of satisfaction with the products"
                    " you received from FNB Business?"
                ),
                (
                    "Q12. Your overall level of satisfaction with FNB Business"
                    " over the last 3 months?"
                ),
                (
                    "Q12.2 Your overall level of satisfaction with your"
                    " Relationship Manager over the last 3-6 months?"
                ),
            ],
            "Total - Mean": [
                fnb_mean,
                bm_mean,
                q10_2_mean,
                q10_3_mean,
                q10_4_mean,
                q10_5_mean,
                q11_1_mean,
                q11_2_mean,
                q11_3_mean,
                q11_4_mean,
                q11_5_mean,
                q12_1_mean,
                q12_2_mean,
                q12_3_mean,
            ],
            "Total - Valid N": [
                fnb_n,
                bm_n,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ],
        }

        summary_df = pd.DataFrame(summary_data)

        st.success("Extraction completed successfully!")

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
            file_name="R10MIL_GROWTH_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
