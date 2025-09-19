import os
import pandas as pd
from sendmail import MailSender
from dotenv import load_dotenv

load_dotenv()

mail_key = os.getenv('MAIL_KEY')

def name_format(name: str) -> str:
    words = name.strip().split()
    result = ' '.join(word.capitalize() for word in words)
    return result

def get_ban_name(sheet_name: str) -> str:
    ban_name = sheet_name.replace('_', ' ').title()
    return ban_name

# Đọc file Excel
filename = 'candidates_gen1.xlsx'
excel_file = pd.ExcelFile(filename)
sheet_names = excel_file.sheet_names

register_dict = {
    'AI Engineer': 'https://forms.gle/PPM82DbhAhsB3S7T9',
    'AI Research': 'https://forms.gle/PPM82DbhAhsB3S7T9',
    'Phần mềm': 'https://forms.gle/PPM82DbhAhsB3S7T9',
    'Nhân sự': 'https://forms.gle/PPM82DbhAhsB3S7T9',
    'Truyền thông': 'https://forms.gle/PPM82DbhAhsB3S7T9',
    'Test': 'https://forms.gle/PPM82DbhAhsB3S7T9',
}

# --- NỘI DUNG EMAIL CHÚC MỪNG ---
html_content_pass = """
<!DOCTYPE html>
<html>
<head>
<title>Thông báo kết quả Vòng đơn PAYT Gen 1</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">

    <p>Xin chào {name},</p>

    <p>Lời đầu tiên, <strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT</strong> gửi lời cảm ơn đến bạn vì đã quan tâm tới đợt tuyển thành viên gen 1 - The Tarot của CLB. Sau thời gian sàng lọc và đánh giá, <strong style="color: #4a4a8a;">PAYT</strong> rất vui khi được thông báo rằng:</p>

    <h3 style="color: #4a4a8a; text-align: center; font-weight: bold;">Bạn đã xuất sắc vượt qua vòng đơn của ban {ban_name}</h3>

    <p style="color: #4a4a8a; text-align: center; font-weight: bold; font-size: 1.1em;">
        🤖 và chính thức tiến vào vòng 2 - Vòng phỏng vấn 🤖
    </p>
    
    <p>Thông tin phỏng vấn cụ thể như sau:</p>

    <ul style="list-style-type: '• '; padding-left: 20px;">
        <li><strong>Thời gian:</strong> Thứ bảy (20/09/2025)</li>
        <li><strong>Địa điểm:</strong> Trung tâm đổi mới sáng tạo IEC Tòa B9, Học viện Công nghệ bưu chính viễn thông.</li>
        <li>Đăng ký lịch phỏng vấn trước 24h00 ngày 18/09 qua link dưới đây:<br><a href="{register_link}" style="color: #1155cc;">{register_link}</a></li>
    </ul>

    <p><strong style="font-style: italic; color: #4a4a8a;">Lưu ý:</strong></p>
    <ul style="list-style-type: '• '; padding-left: 20px;">
        <li>Nếu quá hạn nhưng <strong style="color: #4a4a8a;">PAYT</strong> không nhận được lịch phỏng vấn, <strong style="color: #4a4a8a;">chúng mình</strong> sẽ hiểu rằng bạn không tham gia vòng này.</li>
        <li>Nếu cần thay đổi lịch, hãy liên hệ với chúng mình qua <a href="mailto:paytclub.ptit@gmail.com" style="color: #1155cc;">paytclub.ptit@gmail.com</a> hoặc <a href="https://www.facebook.com/profile.php?id=61579403144221" style="color: #1155cc; text-decoration: underline;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT</a>, tuy nhiên, việc thay đổi lịch ngoài thời gian đã nêu là rất hạn chế.</li>
    </ul>

    <p>Cuối cùng, hãy thật tự tin và chuẩn bị kỹ càng để tham gia vòng phỏng vấn cùng chúng mình. Hành trình đã gần tới hồi kết, <strong style="color: #4a4a8a;">PAYT</strong> đang chờ đợi bạn trở thành mảnh ghép của CLB rồi đó !!</p>

    <p>Trân trọng,</p>

    <p><strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT.</strong></p>

</body>
</html>
"""

# --- NỘI DUNG EMAIL TỪ CHỐI ---
html_content_fail = """
<!DOCTYPE html>
<html>
<head>
<title>Thông báo kết quả Vòng đơn PAYT Gen 1</title>
</head>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
    <p>Xin chào {name},</p>
    <p>Lời đầu tiên, <strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT</strong> gửi lời cảm ơn đến bạn vì đã quan tâm tới đợt tuyển thành viên gen 1 - The Tarot của CLB. Chúng mình rất trân trọng sự quan tâm và nỗ lực của bạn trong quá trình tuyển chọn.</p>
    <p>Sau khi xem xét kỹ lưỡng, rất tiếc chúng mình phải thông báo rằng hồ sơ của bạn chưa phù hợp với các tiêu chí để tiến vào vòng phỏng vấn của {ban_name} trong đợt tuyển chọn lần này. Tuy nhiên, điều này không làm giảm giá trị của những kỹ năng và đam mê mà bạn đã thể hiện.</p>
    <p>Chúng mình khuyến khích bạn tiếp tục trau dồi kiến thức, kỹ năng và tham gia các hoạt động liên quan đến AI. <strong style="color: #4a4a8a;">PAYT</strong> luôn chào đón bạn trong các đợt tuyển chọn tiếp theo hoặc các sự kiện mở mà chúng mình tổ chức sắp tới.</p>
    
    <!-- ===== DÒNG ĐÃ THAY ĐỔI ===== -->
    <p>Nếu bạn có bất kỳ câu hỏi nào hoặc cần phản hồi chi tiết hơn về hồ sơ của mình, vui lòng liên hệ với chúng mình qua email: <a href="mailto:paytclub.ptit@gmail.com" style="color: #1155cc;">paytclub.ptit@gmail.com</a> hoặc <a href="https://www.facebook.com/profile.php?id=61579403144221" style="color: #1155cc; text-decoration: underline;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT</a></p>

    <p>Chúc bạn thành công trong hành trình sắp tới và hy vọng sẽ sớm gặp lại bạn!</p>
    <p>Trân trọng,</p>
    <p><strong style="color: #4a4a8a;">PAYT - Câu lạc bộ trí tuệ nhân tạo PTIT.</strong></p>
</body>
</html>
"""

ourmailsender = MailSender('paytclub.ptit@gmail.com', mail_key, ('smtp.gmail.com', 587))

for sheet_name in ['AI Engineer', 'AI Research', 'Phần mềm', 'Nhân sự', 'Truyền thông']:
    ids = set()
    data = pd.read_excel(filename, sheet_name=sheet_name)
    ban_name = sheet_name
    
    print(f"Dữ liệu từ sheet {sheet_name}:")
    print(data)

    for index, row in data.iterrows():
        if row['Số điện thoại'] in ids:
            continue 
        
        print(f"Processing candidate: {row['Họ và tên']}, Status: {'Pass' if row['Đạt'] else 'Fail'}")
        
        if row['Đạt'] == True:
            personalized_html = html_content_pass.format(
                name=name_format(row['Họ và tên']), 
                ban_name=ban_name, 
                register_link=register_dict.get(ban_name, '')
            )
            plain_backup = f"Xin chào {name_format(row['Họ và tên'])}, Chúc mừng bạn đã vượt qua vòng đơn {ban_name}. Vui lòng mở email trên trình duyệt hoặc ứng dụng hỗ trợ HTML để xem chi tiết."
            
            ourmailsender.set_message(
                in_plaintext=plain_backup,
                in_subject=f"Chúc mừng bạn vượt qua vòng hồ sơ {ban_name}",
                in_from="PAYT - CLB Trí tuệ nhân tạo PTIT",
                in_htmltext=personalized_html
            )
        else:
            personalized_html_fail = html_content_fail.format(
                name=name_format(row['Họ và tên']), 
                ban_name=ban_name
            )
            plain_backup_fail = f"Xin chào {name_format(row['Họ và tên'])}, PAYT xin gửi bạn thông báo về kết quả vòng hồ sơ ban {ban_name}. Vui lòng mở email trên trình duyệt hoặc ứng dụng hỗ trợ HTML để xem chi tiết."

            ourmailsender.set_message(
                in_plaintext=plain_backup_fail,
                in_subject=f"Thông báo kết quả vòng hồ sơ {ban_name}",
                in_from="PAYT - CLB Trí tuệ nhân tạo PTIT",
                in_htmltext=personalized_html_fail
            )
        
        ourmailsender.set_recipients([row['Email của bạn']])
        ourmailsender.connect()
        ourmailsender.send_all()
        
        ids.add(row['Số điện thoại'])