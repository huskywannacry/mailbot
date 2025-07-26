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
    # Định dạng tên ban từ tên sheet (thay "_" thành " " và in hoa chữ cái đầu)
    ban_name = sheet_name.replace('_', ' ').title()
    return ban_name

# Đọc file Excel
filename = 'candidates.xlsx'  # Thay bằng tên file Excel thực tế
excel_file = pd.ExcelFile(filename)
sheet_names = excel_file.sheet_names  # Lấy danh sách tên các sheet

register_dict = {
    'Ban AI Engineer': 'https://docs.google.com/spreadsheets/d/112gYXfqr-FRcdLpdk738rC-tlyrhzBbsbN1plBs5iuM/edit?gid=0#gid=0',
    'Ban AI Research': 'https://docs.google.com/spreadsheets/d/112gYXfqr-FRcdLpdk738rC-tlyrhzBbsbN1plBs5iuM/edit?gid=0#gid=0',
    'Ban Phần mềm': 'https://docs.google.com/spreadsheets/d/1oYfhxRpPqJv3GOY5Nj0PFkqFdsc7Pct15ubMmq1q-pA/edit?usp=sharing',
    'Ban Nhân sự & Truyền thông': 'https://docs.google.com/spreadsheets/d/1bos2QGnfDsA28FsdQjDlbnwIFYDQPcCXSztCI2XBHAo/edit?usp=sharing',
}

# Nội dung email chúc mừng
plaintext_pass = """Xin chào {name},

Chúc mừng bạn đã vượt qua vòng sơ loại hồ sơ của {ban_name} thuộc Câu lạc bộ Tài năng trẻ AI PTIT (PAYT)! Chúng tôi rất ấn tượng với hồ sơ và sự nỗ lực của bạn trong quá trình tuyển chọn.

Để tiếp tục hành trình, mời bạn đăng ký lịch phỏng vấn vào Chủ nhật, ngày 20/07/2025 qua liên kết sau:

{register_link}

Vui lòng hoàn thành đăng ký trước 23:59 ngày 18/07/2025 để đảm bảo tiến độ. Nếu có bất kỳ câu hỏi hoặc cần hỗ trợ, bạn có thể phản hồi lại qua email này.

Chúng tôi rất mong được chào đón bạn vào {ban_name} và cùng nhau tạo nên những giá trị và trải nghiệm tuyệt vời!

Trân trọng,  
Ban Tổ chức PAYT

---------------------
Email: paytclub.ptit@gmail.com
"""

# Nội dung email từ chối
plaintext_fail = """Xin chào {name},

Cảm ơn bạn đã dành thời gian tham gia vòng sơ loại của {ban_name} thuộc Câu lạc bộ Tài năng trẻ AI PTIT (PAYT). Chúng tôi rất trân trọng sự quan tâm và nỗ lực của bạn trong quá trình tuyển chọn.

Sau khi xem xét kỹ lưỡng, rất tiếc chúng tôi phải thông báo rằng hồ sơ của bạn chưa phù hợp với các tiêu chí để tiến vào vòng phỏng vấn của {ban_name} trong đợt tuyển chọn lần này. Tuy nhiên, điều này không làm giảm giá trị của những kỹ năng và đam mê mà bạn đã thể hiện.

Chúng tôi khuyến khích bạn tiếp tục trau dồi kiến thức, kỹ năng và tham gia các hoạt động liên quan đến AI. Câu lạc bộ PAYT luôn chào đón bạn trong các đợt tuyển chọn tiếp theo hoặc các sự kiện mở mà chúng tôi tổ chức.

Nếu bạn có bất kỳ câu hỏi nào hoặc cần phản hồi chi tiết hơn về hồ sơ của mình, vui lòng liên hệ với chúng tôi qua email: paytclub.ptit@gmail.com.

Chúc bạn thành công trong hành trình sắp tới và hy vọng sẽ sớm gặp lại bạn!

Trân trọng,
Ban Tổ chức PAYT
------------------------------

Email: paytclub.ptit@gmail.com
"""

ourmailsender = MailSender('paytclub.ptit@gmail.com', mail_key, ('smtp.gmail.com', 587))

ids = set()

for sheet_name in sheet_names:
    # Đọc dữ liệu từ sheet hiện tại
    data = pd.read_excel(filename, sheet_name=sheet_name)
    ban_name = sheet_name
    
    print(f"Dữ liệu từ sheet {sheet_name}:")
    print(data)

    for index, row in data.iterrows():
        if row['Mã sinh viên:'] in ids:
            continue 

        if row['Trạng thái duyệt'] == 'Pass':
            personalized_message = plaintext_pass.format(name=name_format(row['Họ và tên:']), ban_name=ban_name, register_link=register_dict.get(ban_name, ''))
            ourmailsender.set_message(personalized_message, f"Chúc mừng bạn vượt qua vòng hồ sơ {ban_name}", "Ban Tổ chức PAYT")
        else:
            personalized_message = plaintext_fail.format(name=name_format(row['Họ và tên:']), ban_name=ban_name)
            ourmailsender.set_message(personalized_message, f"Thông báo kết quả vòng hồ sơ {ban_name}", "Ban Tổ chức PAYT")
        
        ourmailsender.set_recipients([row['Email:']])
        ourmailsender.connect()
        ourmailsender.send_all()
        
        ids.add(row['Mã sinh viên:'])