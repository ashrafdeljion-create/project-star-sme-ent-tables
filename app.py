import io
import pandas as pd
import pyreadstat
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# Page configuration
st.set_page_config(
    page_title="SME-ENT Tables Generator",
    page_icon="📊",
    layout="wide",
)

st.title("SME-ENT Tables: R10MIL_GROWTH")
st.write(
    "Upload your SPSS `.sav` data file below. Filters for Type = 1 (Growth)"
    " & 2 (R10Mil), and formats the Excel output with custom FNB brand colors,"
    " full cell borders, and auto-adjusted column widths."
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
      col_map = {c.lower(): c for c in df.columns}

      type_col = col_map.get("type", None)
      seg2_col = col_map.get("seg2", None)
      tmonth_col = col_map.get("tmonth", None)

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

        # Helper: Determine Tmonth values and labels
        month_dict = {}  # code -> label
        if (
            meta
            and meta.variable_value_labels
            and tmonth_col
            and tmonth_col in meta.variable_value_labels
        ):
          month_dict = meta.variable_value_labels[tmonth_col]

        if not month_dict and tmonth_col and tmonth_col in df_tab.columns:
          unique_months = sorted(df_tab[tmonth_col].dropna().unique())
          month_dict = {m: str(m) for m in unique_months}

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

        seg2_filter_col = seg2_col if seg2_col else "seg2"

        # Build columns dynamically
        subsets = {}

        # 1. Total (Overall) Block
        subsets[("Total", "Total", "Total", "Mean")] = df_tab
        subsets[("Total", "Total", "Total", "Valid N")] = df_tab
        subsets[("Total", "Type", "Total", "Mean")] = df_tab
        subsets[("Total", "Type", "Total", "Valid N")] = df_tab
        subsets[("Total", "Type", "Growth", "Mean")] = (
            df_tab[df_tab[type_col] == 1] if type_col else pd.DataFrame()
        )
        subsets[("Total", "Type", "Growth", "Valid N")] = (
            df_tab[df_tab[type_col] == 1] if type_col else pd.DataFrame()
        )
        subsets[("Total", "Type", "R10Mil", "Mean")] = (
            df_tab[df_tab[type_col] == 2] if type_col else pd.DataFrame()
        )
        subsets[("Total", "Type", "R10Mil", "Valid N")] = (
            df_tab[df_tab[type_col] == 2] if type_col else pd.DataFrame()
        )
        subsets[("Total", "Segment", "ENTERPRISE", "Mean")] = (
            df_tab[df_tab[seg2_filter_col] == ent_val]
            if seg2_filter_col in df_tab.columns
            else pd.DataFrame()
        )
        subsets[("Total", "Segment", "ENTERPRISE", "Valid N")] = (
            df_tab[df_tab[seg2_filter_col] == ent_val]
            if seg2_filter_col in df_tab.columns
            else pd.DataFrame()
        )
        subsets[("Total", "Segment", "PLATINUM", "Mean")] = (
            df_tab[df_tab[seg2_filter_col] == plat_val]
            if seg2_filter_col in df_tab.columns
            else pd.DataFrame()
        )
        subsets[("Total", "Segment", "PLATINUM", "Valid N")] = (
            df_tab[df_tab[seg2_filter_col] == plat_val]
            if seg2_filter_col in df_tab.columns
            else pd.DataFrame()
        )

        # 2. Monthly Blocks
        if tmonth_col and month_dict:
          for m_code, m_label in month_dict.items():
            m_df = df_tab[df_tab[tmonth_col] == m_code]
            subsets[(m_label, "Total", "Total", "Mean")] = m_df
            subsets[(m_label, "Total", "Total", "Valid N")] = m_df
            subsets[(m_label, "Type", "Total", "Mean")] = m_df
            subsets[(m_label, "Type", "Total", "Valid N")] = m_df
            subsets[(m_label, "Type", "Growth", "Mean")] = (
                m_df[m_df[type_col] == 1] if type_col else pd.DataFrame()
            )
            subsets[(m_label, "Type", "Growth", "Valid N")] = (
                m_df[m_df[type_col] == 1] if type_col else pd.DataFrame()
            )
            subsets[(m_label, "Type", "R10Mil", "Mean")] = (
                m_df[m_df[type_col] == 2] if type_col else pd.DataFrame()
            )
            subsets[(m_label, "Type", "R10Mil", "Valid N")] = (
                m_df[m_df[type_col] == 2] if type_col else pd.DataFrame()
            )
            subsets[(m_label, "Segment", "ENTERPRISE", "Mean")] = (
                m_df[m_df[seg2_filter_col] == ent_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            )
            subsets[(m_label, "Segment", "ENTERPRISE", "Valid N")] = (
                m_df[m_df[seg2_filter_col] == ent_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            )
            subsets[(m_label, "Segment", "PLATINUM", "Mean")] = (
                m_df[m_df[seg2_filter_col] == plat_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            )
            subsets[(m_label, "Segment", "PLATINUM", "Valid N")] = (
                m_df[m_df[seg2_filter_col] == plat_val]
                if seg2_filter_col in df_tab.columns
                else pd.DataFrame()
            )

        # Define metrics to extract
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
            month_banner, banner, sub_col, stat = col_key
            if is_nps:
              mean_val, n_val = calculate_nps(sub_df, var_name)
              if stat == "Mean":
                row_data[col_key] = mean_val
              else:
                row_data[col_key] = n_val
            else:
              if stat == "Mean":
                row_data[col_key] = calculate_rating_mean(sub_df, var_name)
              else:
                row_data[col_key] = ""
          table_rows.append(row_data)

        # Construct MultiIndex columns dataframe for Streamlit preview
        multi_cols = pd.MultiIndex.from_tuples(
            [("Metric", "", "", "")] + list(subsets.keys()),
            names=["Month", "Banner", "Sub-Group", "Stat"],
        )

        formatted_rows = []
        for r in table_rows:
          flat_row = [r["Metric"]]
          for col_key in subsets.keys():
            flat_row.append(r[col_key])
          formatted_rows.append(flat_row)

        summary_df = pd.DataFrame(formatted_rows, columns=multi_cols)

        st.success("Extraction and monthly breakdown completed successfully!")

        # Display dataframe in app
        st.subheader("Results Preview: R10MIL_GROWTH")
        st.dataframe(summary_df, use_container_width=True)

        # Build custom Excel workbook using openpyxl with FNB logo theme colors
        wb = Workbook()
        ws = wb.active
        ws.title = "R10MIL_GROWTH"
        ws.views.sheetView[0].showGridLines = True

        # Logo Color Palette (Teal/Turquoise, Orange, White, Dark/Black)
        teal_fill = PatternFill(
            start_color="008A90", end_color="008A90", fill_type="solid"
        )  # Primary Brand Teal
        orange_fill = PatternFill(
            start_color="F47920", end_color="F47920", fill_type="solid"
        )  # Brand Orange Accent
        light_teal_fill = PatternFill(
            start_color="E0F2F1", end_color="E0F2F1", fill_type="solid"
        )  # Soft Teal Tint
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

        # Write Headers (Rows 1 to 4)
        col_start = 2
        unique_months = ["Total"] + list(month_dict.values())

        for m_idx, m_name in enumerate(unique_months):
          block_start = col_start
          block_end = col_start + 11

          # Row 1: Month Name
          ws.cell(row=1, column=block_start, value=m_name)
          if block_start != block_end:
            ws.merge_cells(
                start_row=1,
                start_column=block_start,
                end_row=1,
                end_column=block_end,
            )

          # Row 2: Banners
          if m_name == "Total":
            ws.cell(row=2, column=block_start, value="Total")
            ws.merge_cells(
                start_row=2,
                start_column=block_start,
                end_row=2,
                end_column=block_end,
            )
          else:
            ws.cell(row=2, column=block_start, value="Total")
            ws.merge_cells(
                start_row=2,
                start_column=block_start,
                end_row=2,
                end_column=block_start + 1,
            )
            ws.cell(row=2, column=block_start + 2, value="Type")
            ws.merge_cells(
                start_row=2,
                start_column=block_start + 2,
                end_row=2,
                end_column=block_start + 7,
            )
            ws.cell(row=2, column=block_start + 8, value="Seg2")
            ws.merge_cells(
                start_row=2,
                start_column=block_start + 8,
                end_row=2,
                end_column=block_end,
            )

          # Row 3: Sub-Groups
          ws.cell(row=3, column=block_start, value="Total")
          ws.merge_cells(
              start_row=3,
              start_column=block_start,
              end_row=3,
              end_column=block_start + 1,
          )

          ws.cell(row=3, column=block_start + 2, value="Total")
          ws.merge_cells(
              start_row=3,
              start_column=block_start + 2,
              end_row=3,
              end_column=block_start + 3,
          )
          ws.cell(row=3, column=block_start + 4, value="Growth")
          ws.merge_cells(
              start_row=3,
              start_column=block_start + 4,
              end_row=3,
              end_column=block_start + 5,
          )
          ws.cell(row=3, column=block_start + 6, value="R10Mil")
          ws.merge_cells(
              start_row=3,
              start_column=block_start + 6,
              end_row=3,
              end_column=block_start + 7,
          )

          ws.cell(row=3, column=block_start + 8, value="ENTERPRISE")
          ws.merge_cells(
              start_row=3,
              start_column=block_start + 8,
              end_row=3,
              end_column=block_start + 9,
          )
          ws.cell(row=3, column=block_start + 10, value="PLATINUM")
          ws.merge_cells(
              start_row=3,
              start_column=block_start + 10,
              end_row=3,
              end_column=block_end,
          )

          # Row 4: Stats (Mean / Valid N)
          for c in range(block_start, block_end + 1):
            stat_label = "Mean" if c % 2 == 0 else "Valid N"
            ws.cell(row=4, column=c, value=stat_label)

          col_start += 12

        # Insert Data starting at Row 5 with full borders
        for r_idx, row_dict in enumerate(table_rows, start=5):
          # Metric label column (Column A) with border
          metric_cell = ws.cell(row=r_idx, column=1, value=row_dict["Metric"])
          metric_cell.border = data_border
          metric_cell.alignment = Alignment(horizontal="left", vertical="center")

          col_idx = 2
          for col_key in subsets.keys():
            val_cell = ws.cell(
                row=r_idx, column=col_idx, value=row_dict[col_key]
            )
            val_cell.border = data_border
            val_cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            col_idx += 1

        # Apply FNB brand colors and styling to header rows (1 to 4)
        for row in range(1, 5):
          for col in range(1, col_idx):
            cell = ws.cell(row=row, column=col)
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

        # Auto-adjust column widths including Column A to fit all text perfectly
        for col in ws.columns:
          max_len = 0
          col_letter = col[0].column_letter
          for cell in col:
            if cell.value is not None:
              # For column A, account for longer question titles
              val_str = str(cell.value)
              if col_letter == "A":
                max_len = max(max_len, len(val_str))
              else:
                max_len = max(max_len, len(val_str))
          # Set appropriate padding
          if col_letter == "A":
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 30), 65)
          else:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        # Save workbook to BytesIO
        output = io.BytesIO()
        wb.save(output)
        excel_data = output.getvalue()

        # Download button
        st.download_button(
            label="📥 Download Styled Excel Report",
            data=excel_data,
            file_name="R10MIL_GROWTH_Styled_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
