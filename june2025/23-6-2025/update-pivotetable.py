import os
import win32com.client

files_sheets = {
    "bosta.xlsx": ("BO_Delivery", "BO_Delivery Report"),
    "mfb.xlsx": ("MFB_Delivery", "MFB_Delivery Report"),
    "telegraph.xlsx": ("TE_Delivery", "TE_Delivery Report"),
    "vendors.xlsx": ("VE_Delivery", "VE_Delivery Report"),
}

def format_row(row, col_widths):
    return " | ".join(str(cell if cell is not None else "").ljust(col_widths[i]) for i, cell in enumerate(row))

def get_pivot_table_headers(pt):
    try:
        headers = ["Row Labels"]
        data_fields = pt.DataFields
        for i in range(1, data_fields.Count + 1):
            headers.append(data_fields.Item(i).Name)
        return headers
    except Exception:
        return None

def update_pivot_tables():
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False

    grand_totals_sum = []
    grand_totals_headers = []
    max_cols = 0

    try:
        for idx_file, (file, (data_sheet, report_sheet)) in enumerate(files_sheets.items()):
            full_path = os.path.abspath(file)
            if not os.path.exists(full_path):
                print(f"File not found: {full_path}")
                continue

            print(f"\nOpening {file} ...")
            wb = excel.Workbooks.Open(full_path)
            try:
                ws_data = wb.Sheets(data_sheet)
                ws_report = wb.Sheets(report_sheet)

                data_range = ws_data.UsedRange
                data_address = data_range.Address
                print(f"Using source data: '{data_sheet}'!{data_address}")

                pivot_tables = ws_report.PivotTables()
                print(f"pivot_tables count: {pivot_tables.Count}")

                for i in range(1, pivot_tables.Count + 1):
                    pt = pivot_tables.Item(i)
                    print(f"\nUpdating pivot table '{pt.Name}' on sheet '{report_sheet}' ...")
                    pt.ChangePivotCache(
                        wb.PivotCaches().Create(
                            SourceType=win32com.client.constants.xlDatabase,
                            SourceData=f"'{data_sheet}'!{data_address}"
                        )
                    )
                    pt.RefreshTable()

                    values = pt.TableRange2.Value
                    if not values:
                        print("No data found in pivot table.")
                        continue

                    headers = get_pivot_table_headers(pt)
                    if not headers:
                        if isinstance(values[0], tuple):
                            headers = [str(cell if cell is not None else "") for cell in values[0]]
                        else:
                            headers = [str(values[0] if values[0] is not None else "")]

                    # Find Grand Total row
                    grand_total_row = None
                    for row in reversed(values[1:] if len(values) > 1 else []):
                        first_cell = row[0] if len(row) > 0 else None
                        if first_cell and isinstance(first_cell, str) and "grand total" in first_cell.lower():
                            grand_total_row = [cell if cell is not None else "" for cell in row]
                            break

                    if grand_total_row is None:
                        print("Grand Total row not found.")
                        continue

                    # Normalize length to max columns so far
                    current_cols = max(len(headers), len(grand_total_row))
                    if current_cols > max_cols:
                        # Extend sums and headers to new max_cols
                        diff = current_cols - max_cols
                        for _ in range(diff):
                            grand_totals_sum.append(0.0)
                            grand_totals_headers.append("")
                        max_cols = current_cols

                    # Pad headers and grand total row to max_cols
                    headers += [""] * (max_cols - len(headers))
                    grand_total_row += [0] * (max_cols - len(grand_total_row))

                    # Update grand_totals_headers with any new non-empty header
                    for idx_col in range(max_cols):
                        if grand_totals_headers[idx_col] == "" and headers[idx_col] != "":
                            grand_totals_headers[idx_col] = headers[idx_col]

                    col_widths = [max(len(str(grand_totals_headers[i])), len(str(grand_total_row[i]))) for i in range(max_cols)]

                    print(format_row(grand_totals_headers, col_widths))
                    print(format_row(grand_total_row, col_widths))

                    # Sum numeric columns from index 1 onwards (skip 'Row Labels')
                    for idx_col in range(1, max_cols):
                        try:
                            val = float(grand_total_row[idx_col])
                            grand_totals_sum[idx_col] += val
                        except (TypeError, ValueError):
                            pass

                wb.Save()
                print(f"Updated and saved {file}")

                print("\n" + "=" * 80 + "\n")

                if idx_file != len(files_sheets) - 1:
                    print("-" * 80 + "\n")

            except Exception as e:
                print(f"Error processing {file}: {e}")
            finally:
                wb.Close(SaveChanges=False)

    finally:
        excel.Quit()

    if grand_totals_sum and grand_totals_headers:
        print("\n=== Total sums of Grand Totals across all pivot tables ===\n")
        col_widths = [max(len(str(grand_totals_headers[i])), len(str(int(grand_totals_sum[i])) if grand_totals_sum[i].is_integer() else f"{grand_totals_sum[i]:.2f}")) for i in range(max_cols)]
        print(format_row(grand_totals_headers, col_widths))

        total_sum_row = ["Grand Total"]
        for idx_col in range(1, max_cols):
            val = grand_totals_sum[idx_col]
            if isinstance(val, float) and val.is_integer():
                total_sum_row.append(str(int(val)))
            else:
                total_sum_row.append(f"{val:.2f}")
        total_sum_row += [""] * (max_cols - len(total_sum_row))

        print(format_row(total_sum_row, col_widths))

    print("\nAll pivot tables updated!")

if __name__ == "__main__":
    update_pivot_tables()
