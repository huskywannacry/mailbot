import os
import pandas as pd
from sendmail import MailSender
from dotenv import load_dotenv

# Tải các biến môi trường từ file .env
load_dotenv()

# Lấy key gửi mail từ biến môi trường
mail_key = os.getenv('MAIL_KEY')

# Hàm định dạng lại tên (viết hoa chữ cái đầu)
def name_format(name: str) -> str:
    # Chuyển đổi name thành chuỗi để xử lý trường hợp đọc vào là số
    name = str(name)
    words = name.strip().split()
    result = ' '.join(word.capitalize() for word in words)
    return result

# --- NỘI DUNG EMAIL LỊCH PHỎNG VẤN ---
# Mẫu email mới được thiết kế để gửi lịch phỏng vấn
html_content_schedule = """
<!DOCTYPE html>
<html>
<head>
<title>Thông báo Lịch phỏng vấn Vòng 2 - CLB PAYT</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">

    <p>Xin chào {name},</p>

    <p>Một lần nữa, <strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT</strong> xin chúc mừng bạn đã xuất sắc vượt qua Vòng đơn và chính thức tiến vào Vòng phỏng vấn của CLB.</p>
    
    <p>Chúng mình xin gửi bạn thông tin chi tiết về buổi phỏng vấn như sau:</p>

    <div style="background-color: #f4f4f9; border-left: 5px solid #4a4a8a; padding: 15px; margin: 20px 0;">
        <h3 style="color: #4a4a8a; margin-top: 0;">THÔNG TIN PHỎNG VẤN</h3>
        <ul style="list-style-type: '• '; padding-left: 20px;">
            <li><strong>Ban phỏng vấn:</strong> {ban_name}</li>
            <li><strong>Số thứ tự của bạn:</strong> {stt}</li>
            <li><strong>Ca phỏng vấn:</strong> <strong style="color: #d9534f;">{ca_phong_van}</strong></li>
            <li><strong>Thời gian:</strong> Thứ Bảy, ngày 20/09/2025</li>
            <li><strong>Địa điểm:</strong> Trung tâm đổi mới sáng tạo IEC Tòa B9, Học viện Công nghệ bưu chính viễn thông.</li>
        </ul>
    </div>

    <p><strong style="font-style: italic; color: #4a4a8a;">Lưu ý quan trọng:</strong></p>
    <ul style="list-style-type: '• '; padding-left: 20px;">
        <li>Vui lòng có mặt trước ca phỏng vấn của mình từ <strong>5-10 phút</strong> để chuẩn bị tốt nhất.</li>
        <li>Bạn có thể chuẩn bị trước CV (nếu có) hoặc bất kỳ project nào muốn thể hiện trong buổi phỏng vấn.</li>
        <li>Nếu có bất kỳ lý do đột xuất nào không thể tham gia đúng lịch, vui lòng phản hồi lại email này hoặc liên hệ qua fanpage <a href="https://www.facebook.com/profile.php?id=61579403144221" style="color: #1155cc; text-decoration: underline;">PAYT - PTIT AI Club</a> trước 24h00 ngày 19/09/2025.</li>
    </ul>

    <p>Hãy chuẩn bị thật tốt, tự tin và thể hiện hết mình nhé. <strong style="color: #4a4a8a;">PAYT</strong> đang rất mong chờ được gặp và trò chuyện cùng bạn.</p>

    <p>Trân trọng,</p>

    <p><strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT.</strong></p>

</body>
</html>
"""

# Khởi tạo đối tượng gửi mail
# Thay 'your_email@gmail.com' bằng email của CLB
ourmailsender = MailSender('paytclub.ptit@gmail.com', mail_key, ('smtp.gmail.com', 587))

# Đọc file Excel chứa lịch phỏng vấn
filename = 'lich_phong_van.xlsx'
try:
    excel_file = pd.ExcelFile(filename)
    sheet_names = excel_file.sheet_names
except FileNotFoundError:
    print(f"Lỗi: Không tìm thấy file '{filename}'. Vui lòng kiểm tra lại tên file và vị trí.")
    exit()

# Lặp qua từng sheet (tương ứng với mỗi ban)
for sheet_name in sheet_names:
    print(f"\n--- Bắt đầu xử lý cho Ban: {sheet_name} ---")
    data = pd.read_excel(filename, sheet_name=sheet_name)
    
    # Biến để lưu trữ ca phỏng vấn hiện tại
    # Vì trong file excel, ca chỉ được điền ở dòng đầu tiên
    current_ca = ""

    # Lặp qua từng dòng dữ liệu trong sheet
    for index, row in data.iterrows():
        # Bỏ qua các dòng không có Email
        if pd.isna(row['Email']) or not row['Email']:
            continue
            
        # Cập nhật ca phỏng vấn nếu dòng hiện tại có thông tin ca
        if pd.notna(row['Ca']):
            current_ca = row['Ca']
        
        # Lấy thông tin ứng viên từ các cột tương ứng
        # Dùng .get() để tránh lỗi nếu cột không tồn tại
        name = row.get('Tên')
        email = row.get('Email')
        stt = row.get('STT')
        
        # Chuyển đổi STT thành số nguyên để hiển thị đẹp hơn
        try:
            stt = int(stt)
        except (ValueError, TypeError):
            # Nếu STT không phải là số, giữ nguyên giá trị
            pass

        # Điền thông tin cá nhân vào mẫu email
        personalized_html = html_content_schedule.format(
            name=name_format(name), 
            ban_name=sheet_name,
            stt=stt,
            ca_phong_van=current_ca
        )

        # Tạo nội dung backup dạng text (phòng trường hợp email client không hiển thị HTML)
        plain_backup = f"Xin chào {name_format(name)}, PAYT gửi bạn lịch phỏng vấn Vòng 2. Vui lòng mở email trên thiết bị hỗ trợ HTML để xem chi tiết."
        
        # Cấu hình nội dung email
        ourmailsender.set_message(
            in_plaintext=plain_backup,
            in_subject=f"[PAYT] Lịch phỏng vấn Vòng 2 - Ban {sheet_name}",
            in_from="PAYT - CLB Trí tuệ nhân tạo PTIT",
            in_htmltext=personalized_html
        )
        
        # Thiết lập người nhận
        ourmailsender.set_recipients([email])
        
        # Kết nối và gửi mail
        print(f"Đang gửi mail cho: {name_format(name)} ({email}) - STT: {stt} - Ca: {current_ca}")
        try:
            ourmailsender.connect()
            ourmailsender.send_all()
            print(" -> Gửi thành công!")
        except Exception as e:
            print(f" -> Gửi thất bại: {e}")

print("\n--- Hoàn tất quá trình gửi mail! ---")