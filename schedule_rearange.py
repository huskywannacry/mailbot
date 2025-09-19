import pandas as pd
import re
from collections import defaultdict

# --- CẤU HÌNH ---
INPUT_FILE = 'danh_sach_ung_vien.xlsx'
OUTPUT_FILE = 'lich_phong_van_theo_ban.xlsx' # Đổi tên file output cho rõ ràng

# Tên các cột quan trọng
COL_TIMESTAMP = 'Dấu thời gian'
COL_STUDENT_ID = 'Mã sinh viên'
COL_NAME = 'Họ và tên'
COL_DEPARTMENT = 'Ban bạn apply ( với những bạn ứng viên apply 2 ban trở lên, xin vui lòng chọn các ban các bạn apply nhé )'
COL_SCHEDULE_OPTIONS_KEYWORD = 'Lịch phỏng vấn ( hãy chọn ít nhất 2 ca )'

# Danh sách các ca chuẩn và các ban cần loại bỏ
VALID_SHIFTS = ['Ca 1', 'Ca 2', 'Ca 3', 'Ca 4']
EXCLUDED_DEPARTMENTS = ['Nhân sự', 'Truyền thông']

def sanitize_sheet_name(name):
    """Làm sạch tên để phù hợp với quy tắc đặt tên sheet của Excel."""
    # Bỏ các ký tự không hợp lệ
    invalid_chars = r'[]*/\?:\''
    sanitized = re.sub(f'[{re.escape(invalid_chars)}]', '', name)
    # Cắt ngắn nếu dài quá 31 ký tự
    return sanitized[:31]

def find_first_column_with_keyword(df, keyword):
    """Tìm tên cột đầy đủ đầu tiên trong DataFrame chứa từ khóa."""
    for col in df.columns:
        if keyword in col:
            return col
    raise ValueError(f"Không tìm thấy cột nào chứa từ khóa: '{keyword}'")

def schedule_interviews(df, col_department_name, col_schedule_options_name):
    """
    Hàm chính để xử lý và sắp xếp lịch phỏng vấn (Logic giữ nguyên).
    """
    # 1. Ép kiểu và tách các trường hợp "Khác"
    df[col_schedule_options_name] = df[col_schedule_options_name].astype(str)
    is_special_case = df[col_schedule_options_name].str.contains('Khác', na=False)
    special_cases_df = df[is_special_case].copy()
    schedulable_df = df[~is_special_case].copy()
    
    # 2. Trích xuất và đếm số ca rảnh
    schedulable_df['Available_Shifts'] = schedulable_df[col_schedule_options_name].apply(
        lambda x: re.findall(r'Ca \d', str(x))
    )
    schedulable_df['Num_Available'] = schedulable_df['Available_Shifts'].apply(len)
    schedulable_df = schedulable_df[schedulable_df['Num_Available'] > 0].copy()

    # 3. Sắp xếp ưu tiên
    schedulable_df = schedulable_df.sort_values(by='Num_Available', ascending=True)

    # 4. Khởi tạo bộ đếm
    slot_counts = defaultdict(lambda: {shift: 0 for shift in VALID_SHIFTS})
    schedulable_df['Lịch chốt'] = ''

    # 5. Thuật toán xếp lịch
    print("Bắt đầu xếp lịch...")
    for index, row in schedulable_df.iterrows():
        department = row[col_department_name]
        available_shifts = row['Available_Shifts']
        if not department or not available_shifts:
            schedulable_df.loc[index, 'Lịch chốt'] = 'LỖI: Thiếu ban hoặc ca rảnh'
            continue
        possible_options = sorted(
            available_shifts, 
            key=lambda shift: slot_counts[department].get(shift, float('inf'))
        )
        best_shift = possible_options[0] if possible_options else ''
        if best_shift:
            schedulable_df.loc[index, 'Lịch chốt'] = best_shift
            slot_counts[department][best_shift] += 1
        else:
            schedulable_df.loc[index, 'Lịch chốt'] = 'Không tìm được ca phù hợp'
            
    print("Xếp lịch hoàn tất!")
    return schedulable_df, special_cases_df

if __name__ == "__main__":
    try:
        # --- BƯỚC 1 & 2: Đọc và làm sạch dữ liệu (Giữ nguyên) ---
        df = pd.read_excel(INPUT_FILE)
        print(f"Đã đọc thành công {len(df)} dòng từ file '{INPUT_FILE}'.")
        
        df = df.sort_values(by=COL_TIMESTAMP, ascending=False)
        df = df.drop_duplicates(subset=[COL_STUDENT_ID], keep='first')
        print(f"Đã loại bỏ đơn đăng ký trùng lặp, còn lại {len(df)} đơn.")

        # --- BƯỚC 3: Lọc bỏ các ban không cần xếp lịch (Giữ nguyên) ---
        exclusion_pattern = '|'.join(EXCLUDED_DEPARTMENTS)
        mask = ~df[COL_DEPARTMENT].str.contains(exclusion_pattern, na=False, case=False)
        df_filtered = df[mask].copy()
        print(f"Đã lọc bỏ ban Nhân sự/Truyền thông, còn lại {len(df_filtered)} ứng viên.")

        # --- BƯỚC 4: Tiến hành xếp lịch (Giữ nguyên) ---
        if not df_filtered.empty:
            col_schedule_options_actual = find_first_column_with_keyword(df_filtered, COL_SCHEDULE_OPTIONS_KEYWORD)
            
            scheduled, special_cases = schedule_interviews(df_filtered, COL_DEPARTMENT, col_schedule_options_actual)

            if scheduled is not None and not scheduled.empty:
                scheduled = scheduled.drop(columns=['Available_Shifts', 'Num_Available'], errors='ignore')

                # --- BƯỚC 5: LƯU KẾT QUẢ THEO TỪNG BAN (PHẦN THAY ĐỔI) ---
                print("Đang xuất kết quả ra file Excel, mỗi ban một sheet...")
                with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
                    # 1. Ghi sheet các trường hợp đặc biệt trước
                    if not special_cases.empty:
                        special_cases.to_excel(writer, sheet_name='Cần Xếp Lịch Riêng', index=False)
                    
                    # 2. Nhóm kết quả đã xếp theo từng ban
                    grouped_by_dept = scheduled.groupby(COL_DEPARTMENT)
                    
                    # 3. Lặp qua mỗi ban và ghi ra một sheet riêng
                    for department_name, department_df in grouped_by_dept:
                        # Làm sạch tên ban để dùng làm tên sheet
                        sheet_name = sanitize_sheet_name(department_name)
                        
                        # Sắp xếp theo ca chốt để dễ nhìn
                        department_df_sorted = department_df.sort_values(by='Lịch chốt')
                        
                        department_df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"\n✅ THÀNH CÔNG! Kết quả đã được lưu vào file '{OUTPUT_FILE}'")
                print("File kết quả chứa các sheet tương ứng với từng ban và 1 sheet 'Cần Xếp Lịch Riêng'.")
            else:
                print("\nKhông có ứng viên nào hợp lệ để xếp lịch.")
        else:
            print("\nKhông có ứng viên nào hợp lệ để xếp lịch sau khi lọc.")

    except Exception as e:
        print(f"\n❌ Đã xảy ra lỗi không mong muốn: {e}")
        import traceback
        traceback.print_exc()