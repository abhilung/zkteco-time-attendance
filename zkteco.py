# ZKteco attendance machine report

import os
import time
from zk import ZK, const
import datetime
import nepali_datetime
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from collections import defaultdict

def get_save_directory():
    home = os.path.expanduser("~")

    possible_paths = [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "Downloads"),
    ]

    for path in possible_paths:
        if os.path.isdir(path):
            return path

    # Last resort: create Downloads
    fallback = os.path.join(home, "Downloads")
    os.makedirs(fallback, exist_ok=True)
    return fallback

SAVE_DIR = get_save_directory()

# -------------------------------
# RETRY HELPER (Silent retry)
# -------------------------------
def retry_operation(func, retries=3, delay=3, operation_name="operation"):
    for attempt in range(1, retries + 1):
        try:
            return func()
        except Exception as e:
            print(f"{operation_name} failed with error: {e}")
            if attempt < retries:
                print(f"{operation_name} failed. Retrying ({attempt}/{retries})...")
                time.sleep(delay)
            else:
                print(f"{operation_name} failed after {retries} attempts.")
                raise

def get_unique_filename(folder, base_name):
    file_path = os.path.join(folder, f"{base_name}.xlsx")
    if not os.path.exists(file_path):
        return file_path

    counter = 2
    while True:
        file_path = os.path.join(folder, f"{base_name}{counter}.xlsx")
        if not os.path.exists(file_path):
            return file_path
        counter += 1

def export_all_staff_to_excel():
    ip_address = '192.168.1.201'  # This is default ip address from company for the machine
    port = 4370
    zk = ZK(ip_address, port=port, timeout=5, password=0, force_udp=False)
    conn = None

    try:
        # -------------------------------
        # CONNECT TO MACHINE
        # -------------------------------
        print("Connecting to device...")
        conn = retry_operation(
            zk.connect,
            retries=3,
            delay=3,
            operation_name="Device connection",
        )
        print("Connected successfully.")

        # -------------------------------
        # WELCOME & MACHINE DETAILS
        # -------------------------------
        print("\n================================")
        print(" Welcome to ZKTeco Attendance Monthly Report Export Program!!")
        print("================================\n")

        print("--- MACHINE DETAILS ---")
        print("Device Name     :", conn.get_device_name())
        print("Serial Number   :", conn.get_serialnumber())
        print("Firmware Ver.   :", conn.get_firmware_version())
        print("Platform        :", conn.get_platform())

        # -------------------------------
        # USERS & LOG COUNT
        # -------------------------------
        users = retry_operation(
            conn.get_users,
            retries=3,
            delay=2,
            operation_name="Fetching users",
        )

        attendance = retry_operation(
            conn.get_attendance,
            retries=3,
            delay=2,
            operation_name="Fetching attendance logs",
        )

        attendance_by_user = defaultdict(list)
                
        for att in attendance:
            attendance_by_user[att.user_id].append(att)

        print(f"\n--- USERS ({len(users)}) ---")
        for user in users:
            print(f"User ID: {user.user_id} | Name: {user.name}")

        print(f"\nTotal Attendance Logs Stored: {len(attendance)}/200000")
        print("--------------------------------\n")

        # -------------------------------
        # ASK MONTH
        # -------------------------------
        try:
            target_month_nep = int(input("Enter Nepali Month number (1-12): "))
            if not 1 <= target_month_nep <= 12:
                print("Invalid month. Must be between 1 and 12.")
                return
        except ValueError:
            print("Invalid input. Month must be a number.")
            return

        today_nepali = nepali_datetime.datetime.now()
        target_year_nep = today_nepali.year
        month_name = nepali_datetime.date(
            target_year_nep, target_month_nep, 1
        ).strftime('%B')

        print(f"Processing attendance for {month_name} {target_year_nep} BS...\n")

        # Disable device AFTER info & input
        conn.disable_device()

        # -------------------------------
        # LAST DAY OF NEPALI MONTH
        # -------------------------------
        last_day_nep = 32
        while last_day_nep > 28:
            try:
                nepali_datetime.date(target_year_nep, target_month_nep, last_day_nep)
                break
            except ValueError:
                last_day_nep -= 1

        base_file_name = f"{month_name}_attendance_report"
        file_name = get_unique_filename(SAVE_DIR, base_file_name)
        
        # -------------------------------
        # EXCEL EXPORT
        # -------------------------------
        print(f"Saving report to: {SAVE_DIR}")
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for user in users:
                user_attendance_map = {}

                for att in attendance_by_user.get(user.user_id, []):
                    date_bs = nepali_datetime.date.from_datetime_date(
                        att.timestamp.date()
                    )

                    if (
                        date_bs.year == target_year_nep
                        and date_bs.month == target_month_nep
                    ):
                        user_attendance_map.setdefault(date_bs.day, {
                            "in": [],
                            "out": []
                        })

                        if hasattr(att, "punch"):
                            if att.punch == 0:
                                user_attendance_map[date_bs.day]["in"].append(att.timestamp)
                            elif att.punch == 1:
                                user_attendance_map[date_bs.day]["out"].append(att.timestamp)

                sheet_name = "".join(
                    x for x in user.name if x.isalnum() or x == " "
                )[:31].strip()
                if not sheet_name:
                    sheet_name = f"User_{user.user_id}"

                ws = writer.book.create_sheet(sheet_name)

                # -------------------------------
                # STYLES
                # -------------------------------
                bold_font = Font(bold=True)
                off_font = Font(bold=True, color="006100")
                off_fill = PatternFill(
                    start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"
                )
                center_align = Alignment(horizontal="center")
                left_align = Alignment(horizontal="left")

                TABLE_HEADER_ROW = 1

                headers = ["BS Date", "Day", "Clock In", "Clock Out", "Worked Hours", "Comments"]
                for col, header in zip(["D","E","F","G","H","I"], headers):
                    ws[f"{col}{TABLE_HEADER_ROW}"] = header
                    ws[f"{col}{TABLE_HEADER_ROW}"].font = bold_font

                # -------------------------------
                # LEFT INFO BLOCK
                # -------------------------------
                INFO_START_ROW = 8

                ws[f"A{INFO_START_ROW}"] = "Staff Name:"
                ws[f"B{INFO_START_ROW}"] = user.name

                ws[f"A{INFO_START_ROW+1}"] = "Staff ID:"
                ws[f"B{INFO_START_ROW+1}"] = user.user_id

                ws[f"A{INFO_START_ROW+2}"] = "Report Month:"
                ws[f"B{INFO_START_ROW+2}"] = f"{month_name} {target_year_nep} BS"

                for r in range(INFO_START_ROW, INFO_START_ROW + 3):
                    ws[f"A{r}"].font = bold_font
                    ws[f"B{r}"].font = bold_font
                    
                # -------------------------------
                # DATA ROWS (Revised for Interactivity)
                # -------------------------------
                for day in range(1, last_day_nep + 1):
                    current_row = TABLE_HEADER_ROW + day
                    date_label = f"{target_year_nep}-{target_month_nep:02d}-{day:02d}"

                    day_name = nepali_datetime.date(
                        target_year_nep, target_month_nep, day
                    ).strftime("%A")

                    # This formula is the "brain". It checks if F and G are numbers.
                    # If you delete "OFF" and type a time, it calculates immediately.
                    row_formula = (
                        f'=IF(AND(ISNUMBER(F{current_row}), ISNUMBER(G{current_row})), '
                        f'ROUND(MAX(0, G{current_row}-F{current_row})*1440,0)/1440, 0)'
                    )

                    ws[f"D{current_row}"] = date_label
                    ws[f"E{current_row}"] = day_name
                    ws[f"H{current_row}"] = row_formula # Apply formula to every single row
                    ws[f"H{current_row}"].number_format = "h:mm"

                    is_future = (
                        (target_year_nep, target_month_nep, day)
                        > (today_nepali.year, today_nepali.month, today_nepali.day)
                    )

                    if is_future:
                        # Leave In/Out blank for future dates
                        pass 

                    elif day in user_attendance_map:
                        day_logs = user_attendance_map[day]
                        in_times = sorted(day_logs["in"])
                        out_times = sorted(day_logs["out"])

                        in_time = in_times[0].time() if in_times else None
                        out_time = out_times[-1].time() if out_times else None

                        ws[f"F{current_row}"] = in_time
                        ws[f"G{current_row}"] = out_time
                        ws[f"F{current_row}"].number_format = "h:mm AM/PM"
                        ws[f"G{current_row}"].number_format = "h:mm AM/PM"

                    else:
                        # For days with no logs, we put "OFF"
                        ws[f"F{current_row}"] = "OFF"
                        ws[f"G{current_row}"] = "OFF"
                        
                        # Apply the green styling
                        for col in ["F","G"]:
                            ws[f"{col}{current_row}"].font = off_font
                            ws[f"{col}{current_row}"].fill = off_fill

                    ws[f"I{current_row}"] = ""


                # -------------------------------
                # SUMMARY
                # -------------------------------
                total_row_idx = TABLE_HEADER_ROW + last_day_nep + 1

                ws[f"D{total_row_idx}"] = "TOTAL WORKED HOURS:"
                ws[f"D{total_row_idx}"].font = bold_font

                ws[f"H{total_row_idx}"] = f"=SUM(H2:H{total_row_idx-1})"
                ws[f"H{total_row_idx}"].number_format = '[h]:mm "HRS"'
                ws[f"H{total_row_idx}"].font = bold_font

                # Mirror total worked hours to left summary (A11)
                ws[f"A{INFO_START_ROW+3}"] = "Total Monthly Hours:"
                ws[f"B{INFO_START_ROW+3}"] = f"=H{total_row_idx}"
                ws[f"B{INFO_START_ROW+3}"].number_format = '[h]:mm "HRS"'

                ws[f"A{INFO_START_ROW+3}"].font = bold_font
                ws[f"B{INFO_START_ROW+3}"].font = bold_font


                # -------------------------------
                # FORMATTING
                # -------------------------------
                
                ws.column_dimensions["A"].width = 24   # Column 'A'
                ws.column_dimensions["B"].width = 22   # Column 'B'
                ws.column_dimensions["D"].width = 15   # BS Date
                ws.column_dimensions["E"].width = 15   # Day
                ws.column_dimensions["F"].width = 20   # Clock In
                ws.column_dimensions["G"].width = 20   # Clock Out
                ws.column_dimensions["H"].width = 20   # Worked Hours
                ws.column_dimensions["I"].width = 26   # Comments / Notes

                for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                    for cell in row:
                        cell.alignment = (
                            left_align if cell.column_letter == "E"
                            else center_align
                        )

        if "Sheet" in writer.book.sheetnames:
            del writer.book["Sheet"]

        print(f"\nReport saved: {file_name}")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        if conn:
            conn.enable_device()
            conn.disconnect()

if __name__ == "__main__":
    export_all_staff_to_excel()
    input("\nPress Enter to exit...")
