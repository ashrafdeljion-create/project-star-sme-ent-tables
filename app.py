import io
import pandas as pd
import pyreadstat
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Page configuration
st.set_page_config(
    page_title="SME-ENT Tables Generator",
    page_icon="📊",
    layout="wide",
)

st.title("SME-ENT Tables Generator: R10MIL_GROWTH & PUBSC")
st.write(
    "Upload your SPSS `.sav` data file below. Generates multi-sheet workbook"
    " for R10MIL_GROWTH (Type 1 & 2) and PUBSC (Type 3 with SUBREG banner)."
)

# File uploader widget
uploaded_file = st.file_uploader(
    "Upload .SAV File", type=["sav"], help="Upload your SPSS dataset here."
)

# Execution button
if uploaded_file is not None:
  if st.button("Run Extraction"):
    with st.spinner("Processing dataset and generating workbook..."):
      # Save uploaded file temporarily so pyreadstat can read it
      with open("temp.sav", "wb") as f:
        f.write(uploaded_file.getbuffer())

      # Read the .sav file and metadata using pyreadstat
      df, meta = pyreadstat.read_sav("temp.sav")

      # Normalize column names for case-insensitive matching
      col_map = {c.lower(): c for c in df.columns}

      type_col = col_map.get("type", None)
      seg2_col = col_map.get("seg2", None)
      tmonth_col = col_map.get("tmonth", None)
      subreg_col = col_map.get("subreg", None)

      if not type_col:
        st.error(
            "Error: 'Type' column not found in the uploaded SPSS file."
            " Please check your variable names."
        )
      else:

        # Helper: Determine segment codes dynamically from metadata
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

        # Helper: Determine Tmonth values and labels
        month_dict = {}
        if (
            meta
            and meta.variable_value_labels
            and tmonth_col
            and tmonth_col in meta.variable_value_labels
        ):
          month_dict = meta.variable_value_labels[tmonth_col]
        if not month_dict and tmonth_col and tmonth_col in df.columns:
          unique_months = sorted(df[tmonth_col].dropna().unique())
          month_dict = {m: str(m) for m in unique_months}

        # Helper: Determine SUBREG values and labels for PUBSC
        subreg_dict = {}
        if (
            meta
            and meta.variable_value_labels
            and subreg_col
            and subreg_col in meta.variable_value_labels
        ):
          subreg_dict = meta.variable_value_labels[subreg_col]
        if not subreg_dict and subreg_col and subreg_col in df.columns:
          unique_subs = sorted(df[subreg_col].dropna().unique())
          subreg_dict = {s: str(s) for s in unique_subs}

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

        # Metrics configuration list shared across sheets
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

        # ==========================================
        # DATASET 1: R10MIL_GROWTH (Type 1 & 2)
        # ==========================================
        df_r10_growth = df[df[type_col].isin([1, 2])]
        seg2_filter_col = seg2_col if seg2_col else "seg2"

        subsets_r10 = {}
        # Total overall block
        subsets_r10[("Total", "Total", "Total", "Mean")] = df_r10_growth
        subsets_r10[("Total", "Total", "Total", "Valid N")] = df_r10_growth
        subsets_r10[("Total", "Type", "Total", "Mean")] = df_r10_growth
        subsets_r10[("Total", "Type", "Total", "Valid N")] = df_r10_growth
        subsets_r10[("Total", "Type", "Growth", "Mean")] = df_r10_growth[
            df_r10_growth[type_col] == 1
        ]
        subsets_r10[("Total", "Type", "Growth", "Valid N")] = df_r10_growth[
            df_r10_growth[type_col] == 1
        ]
        subsets_r10[("Total", "Type", "R10Mil", "Mean")] = df_r10_growth[
            df_r10_growth[type_col] == 2
        ]
        subsets_r10[("Total", "Type", "R10Mil", "Valid N")] = df_r10_growth[
            df_r10_growth[type_col] == 2
        ]
        subsets_r10[("Total", "Segment", "ENTERPRISE", "Mean")] = (
            df_r10_growth[df_r10_growth[seg2_filter_col] == ent_val]
            if seg2_filter_col in df_r10_growth.columns
            else pd.DataFrame()
        )
        subsets_r10[("Total", "Segment", "ENTERPRISE", "Valid N")] = (
            df_r10_growth[df_r10_growth[seg2_filter_col] == ent_val]
            if seg2_filter_col in df_r10_growth.columns
            else pd.DataFrame()
        )
        subsets_r10[("Total", "Segment", "PLATINUM", "Mean")] = (
            df_r10_growth[df_r10_growth[seg2_filter_col] == plat_val]
            if seg2_filter_col in df_r10_growth.columns
            else pd.DataFrame()
        )
        subsets_r10[("Total", "Segment", "PLATINUM", "Valid N")] = (
            df_r10_growth[df_r10_growth[seg2_filter_col] == plat_val]
            if seg2_filter_col in df_r10_growth.columns
            else pd.DataFrame()
        )

        # Monthly blocks for R10MIL_GROWTH
        if tmonth_col and month_dict:
          for m_code, m_label in month_dict.items():
            m_df = df_r10_growth[df_r10_growth[tmonth_col] == m_code]
            subsets_r10[(m_label, "Total", "Total", "Mean")] = m_df
            subsets_r10[(m_label, "Total", "Total", "Valid N")] = m_df
            subsets_r10[(m_label, "Type", "Total", "Mean")] = m_df
            subsets_r10[(m_label, "Type", "Total", "Valid N")] = m_df
            subsets_r10[(m_label, "Type", "Growth", "Mean")] = m_df[
                m_df[type_col] == 1
            ]
            subsets_r10[(m_label, "Type", "Growth", "Valid N")] = m_df[
                m_df[type_col] == 1
            ]
            subsets_r10[(m_label, "Type", "R10Mil", "Mean")] = m_df[
                m_df[type_col] == 2
            ]
            subsets_r10[(m_label, "Type", "R10Mil", "Valid N")] = m_df[
                m_df[type_col] == 2
            ]
            subsets_r10[(m_label, "Segment", "ENTERPRISE", "Mean")] = (
                m_df[m_df[seg2_filter_col] == ent_val]
                if seg2_filter_col in m_df.columns
                else pd.DataFrame()
            )
            subsets_r10[(m_label, "Segment", "ENTERPRISE", "Valid N")] = (
                m_df[m_df[seg2_filter_col] == ent_val]
                if seg2_filter_col in m_df.columns
                else pd.DataFrame()
            )
            subsets_r10[(m_label, "Segment", "PLATINUM", "Mean")] = (
                m_df[m_df[seg2_filter_col] == plat_val]
                if seg2_filter_col in m_df.columns
                else pd.DataFrame()
            )
            subsets_r10[(m_label, "Segment", "PLATINUM", "Valid N")] = (
                m_df[m_df[seg2_filter_col] == plat_val]
                if seg2_filter_col in m_df.columns
                else pd.DataFrame()
            )

        # Compute rows for R10MIL_GROWTH
        rows_r10 = []
        for label, var_name, is_nps in metrics_config:
          row_data = {"Metric": label}
          for col_key, sub_df in subsets_r10.items():
            _, _, _, stat = col_key
            if is_nps:
              mean_val, n_val = calculate_nps(sub_df, var_name)
              row_data[col_key] = mean_val if stat == "Mean" else n_val
            else:
              row_data[col_key] = (
                  calculate_rating_mean(sub_df, var_name)
                  if stat == "Mean"
                  else ""
              )
          rows_r10.append(row_data)

        # ==========================================
        # DATASET 2: PUBSC (Type 3 with SUBREG banner)
        # ==========================================
        df_pubsc = df[df[type_col] == 3]
        subreg_filter_col = subreg_col if subreg_col else "subreg"

        subsets_pub = {}
        # Total overall block for PUBSC
        subsets_pub[("Total", "Total", "Mean")] = df_pubsc
        subsets_pub[("Total", "Total", "Valid N")] = df_pubsc
        if subreg_filter_col in df_pubsc.columns:
          for s_code, s_label in subreg_dict.items():
            s_df = df_pubsc[df_pubsc[subreg_filter_col] == s_code]
            subsets_pub[("Total", s_label, "Mean")] = s_df
            subsets_pub[("Total", s_label, "Valid N")] = s_df

        # Monthly blocks for PUBSC with SUBREG sub-banners
        if tmonth_col and month_dict:
          for m_code, m_label in month_dict.items():
            m_df = df_pubsc[df_pubsc[tmonth_col] == m_code]
            subsets_pub[(m_label, "Total", "Mean")] = m_df
            subsets_pub[(m_label, "Total", "Valid N")] = m_df
            if subreg_filter_col in m_df.columns:
              for s_code, s_label in subreg_dict.items():
                s_df = m_df[m_df[subreg_filter_col] == s_code]
                subsets_pub[(m_label, s_label, "Mean")] = s_df
                subsets_pub[(m_label, s_label, "Valid N")] = s_df

        # Compute rows for PUBSC
        rows_pub = []
        for label, var_name, is_nps in metrics_config:
          row_data = {"Metric": label}
          for col_key, sub_df in subsets_pub.items():
            # col_key is either (month, 'Total', stat) or (month, region, stat)
            stat = col_key[-1]
            if is_nps:
              mean_val, n_val = calculate_nps(sub_df, var_name)
              row_data[col_key] = mean_val if stat == "Mean" else n_val
            else:
              row_data[col_key] = (
                  calculate_rating_mean(sub_df, var_name)
                  if stat == "Mean"
                  else ""
              )
          rows_pub.append(row_data)

        st.success(
            "Extraction completed successfully for both R10MIL_GROWTH and"
            " PUBSC!"
        )

        # Preview in Streamlit
        st.subheader("Results Preview: R10MIL_GROWTH")
        st.dataframe(pd.DataFrame(rows_r10), use_container_width=True)

        st.subheader("Results Preview: PUBSC")
        st.dataframe(pd.DataFrame(rows_pub), use_container_width=True)

        # ==========================================
        # BUILD EXCEL WORKBOOK (Multi-Sheet)
        # ==========================================
        wb = Workbook()
        # Remove default sheet
        default_sheet = wb.active

        # Logo Color Palette (Teal, Orange, Light Teal, Grays)
        teal_fill = PatternFill(
            start_color="008A90", end_color="008A90", fill_type="solid"
        )
        orange_fill = PatternFill(
            start_color="F47920", end_color="F47920", fill_type="solid"
        )
        light_teal_fill = PatternFill(
            start_color="E0F2F1", end_color="E0F2F1", fill_type="solid"
        )
        white_font = Font(color="FFFFFF", bold=True, size=10)
        dark_font = Font(color="000000", bold=True, size=10)
        thin_border = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )
        data_border = Border(
            left=Side(style="thin", color="E0E0E0"),
            right=Side(style="thin", color="E0E0E0"),
            top=Side(style="thin", color="E0E0E0"),
            bottom=Side(style="thin", color="E0E0E0"),
        )

        # --- SHEET 1: R10MIL_GROWTH ---
        ws1 = wb.create_sheet(title="R10MIL_GROWTH")
        ws1.views.sheetView[0].showGridLines = True

        col_start = 2
        unique_months = ["Total"] + list(month_dict.values())

        for m_name in unique_months:
          block_start = col_start
          block_end = col_start + 11

          ws1.cell(row=1, column=block_start, value=m_name)
          if block_start != block_end:
            ws1.merge_cells(
                start_row=1,
                start_column=block_start,
                end_row=1,
                end_column=block_end,
            )

          if m_name == "Total":
            ws1.cell(row=2, column=block_start, value="Total")
            ws1.merge_cells(
                start_row=2,
                start_column=block_start,
                end_row=2,
                end_column=block_end,
            )
          else:
            ws1.cell(row=2, column=block_start, value="Total")
            ws1.merge_cells(
                start_row=2,
                start_column=block_start,
                end_row=2,
                end_column=block_start + 1,
            )
            ws1.cell(row=2, column=block_start + 2, value="Type")
            ws1.merge_cells(
                start_row=2,
                start_column=block_start + 2,
                end_row=2,
                end_column=block_start + 7,
            )
            ws1.cell(row=2, column=block_start + 8, value="Seg2")
            ws1.merge_cells(
                start_row=2,
                start_column=block_start + 8,
                end_row=2,
                end_column=block_end,
            )

          ws1.cell(row=3, column=block_start, value="Total")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start,
              end_row=3,
              end_column=block_start + 1,
          )
          ws1.cell(row=3, column=block_start + 2, value="Total")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start + 2,
              end_row=3,
              end_column=block_start + 3,
          )
          ws1.cell(row=3, column=block_start + 4, value="Growth")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start + 4,
              end_row=3,
              end_column=block_start + 5,
          )
          ws1.cell(row=3, column=block_start + 6, value="R10Mil")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start + 6,
              end_row=3,
              end_column=block_start + 7,
          )
          ws1.cell(row=3, column=block_start + 8, value="ENTERPRISE")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start + 8,
              end_row=3,
              end_column=block_start + 9,
          )
          ws1.cell(row=3, column=block_start + 10, value="PLATINUM")
          ws1.merge_cells(
              start_row=3,
              start_column=block_start + 10,
              end_row=3,
              end_column=block_end,
          )

          for c in range(block_start, block_end + 1):
            ws1.cell(
                row=4, column=c, value="Mean" if c % 2 == 0 else "Valid N"
            )
          col_start += 12

        max_col_r10 = col_start - 1
        for r_idx, row_dict in enumerate(rows_r10, start=5):
          metric_cell = ws1.cell(row=r_idx, column=1, value=row_dict["Metric"])
          metric_cell.border = data_border
          metric_cell.alignment = Alignment(horizontal="left", vertical="center")
          col_idx = 2
          for col_key in subsets_r10.keys():
            val_cell = ws1.cell(
                row=r_idx, column=col_idx, value=row_dict[col_key]
            )
            val_cell.border = data_border
            val_cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            col_idx += 1

        for row in range(1, 5):
          for col in range(1, max_col_r10 + 1):
            cell = ws1.cell(row=row, column=col)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = thin_border
            if row == 1:
              cell.fill = teal_fill
              cell.font = white_font
            elif row == 2:
              cell.fill = orange_fill
              cell.font = white_font
            elif row == 3:
              cell.fill = light_teal_fill
              cell.font = dark_font
            else:
              cell.fill = PatternFill(
                  start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"
              )
              cell.font = dark_font

        for col_num in range(1, max_col_r10 + 1):
          col_letter = get_column_letter(col_num)
          max_len = 0
          for row_num in range(1, len(rows_r10) + 6):
            cell_val = ws1.cell(row=row_num, column=col_num).value
            if cell_val is not None:
              max_len = max(max_len, len(str(cell_val)))
          ws1.column_dimensions[col_letter].width = (
              min(max(max_len + 4, 30), 65)
              if col_letter == "A"
              else max(max_len + 3, 12)
          )

        # --- SHEET 2: PUBSC ---
        ws2 = wb.create_sheet(title="PUBSC")
        ws2.views.sheetView[0].showGridLines = True

        # Each month block in PUBSC has 2 columns for Total (Mean/Valid N) + 2 columns per region in subreg_dict
        regions_count = len(subreg_dict) if subreg_dict else 0
        cols_per_month = 2 + (regions_count * 2)

        col_start_p = 2
        for m_name in unique_months:
          block_start = col_start_p
          block_end = col_start_p + cols_per_month - 1

          # Row 1: Month Name
          ws2.cell(row=1, column=block_start, value=m_name)
          if block_start != block_end:
            ws2.merge_cells(
                start_row=1,
                start_column=block_start,
                end_row=1,
                end_column=block_end,
            )

          # Row 2: Banner ("SUBREG")[cite: 18]
          ws2.cell(row=2, column=block_start, value="SUBREG")
          ws2.merge_cells(
              start_row=2,
              start_column=block_start,
              end_row=2,
              end_column=block_end,
          )

          # Row 3: Sub-Groups (Total + Each Region)[cite: 18]
          ws2.cell(row=3, column=block_start, value="Total")
          ws2.merge_cells(
              start_row=3,
              start_column=block_start,
              end_row=3,
              end_column=block_start + 1,
          )

          curr_c = block_start + 2
          if subreg_dict:
            for s_code, s_label in subreg_dict.items():
              ws2.cell(row=3, column=curr_c, value=s_label)
              ws2.merge_cells(
                  start_row=3,
                  start_column=curr_c,
                  end_row=3,
                  end_column=curr_c + 1,
              )
              curr_c += 2

          # Row 4: Stats (Mean / Valid N)
          for c in range(block_start, block_end + 1):
            ws2.cell(
                row=4, column=c, value="Mean" if c % 2 == 0 else "Valid N"
            )

          col_start_p += cols_per_month

        max_col_pub = col_start_p - 1
        for r_idx, row_dict in enumerate(rows_pub, start=5):
          metric_cell = ws2.cell(row=r_idx, column=1, value=row_dict["Metric"])
          metric_cell.border = data_border
          metric_cell.alignment = Alignment(horizontal="left", vertical="center")
          col_idx = 2
          for col_key in subsets_pub.keys():
            val_cell = ws2.cell(
                row=r_idx, column=col_idx, value=row_dict[col_key]
            )
            val_cell.border = data_border
            val_cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            col_idx += 1

        for row in range(1, 5):
          for col in range(1, max_col_pub + 1):
            cell = ws2.cell(row=row, column=col)
            cell.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell.border = thin_border
            if row == 1:
              cell.fill = teal_fill
              cell.font = white_font
            elif row == 2:
              cell.fill = orange_fill
              cell.font = white_font
            elif row == 3:
              cell.fill = light_teal_fill
              cell.font = dark_font
            else:
              cell.fill = PatternFill(
                  start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"
              )
              cell.font = dark_font

        for col_num in range(1, max_col_pub + 1):
          col_letter = get_column_letter(col_num)
          max_len = 0
          for row_num in range(1, len(rows_pub) + 6):
            cell_val = ws2.cell(row=row_num, column=col_num).value
            if cell_val is not None:
              max_len = max(max_len, len(str(cell_val)))
          ws2.column_dimensions[col_letter].width = (
              min(max(max_len + 4, 30), 65)
              if col_letter == "A"
              else max(max_len + 3, 12)
          )

        # Remove default blank sheet if present
        if default_sheet in wb.worksheets:
          wb.remove(default_sheet)

        # Save workbook to BytesIO
        output = io.BytesIO()
        wb.save(output)
        excel_data = output.getvalue()

        # Download button
        st.download_button(
            label="📥 Download Complete Multi-Sheet Excel Report",
            data=excel_data,
            file_name="SME_ENT_and_PUBSC_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
