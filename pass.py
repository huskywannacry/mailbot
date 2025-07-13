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

# Đọc danh sách email từ file Excel
data = pd.read_excel('answer.xlsx')
# mails = data.Email.tolist()

# Nội dung email
plaintext = """Xin chào {name},

Chúc mừng bạn đã vượt qua vòng sơ loại đầu tiên của Câu lạc bộ Tài năng trẻ AI PTIT (PAYT)! Chúng tôi rất ấn tượng với hồ sơ và sự hiểu biết của bạn trong quá trình tuyển chọn.

Để tiếp tục hành trình, mời bạn hoàn thành việc nộp CV qua biểu mẫu sau:

https://docs.google.com/forms/d/e/1FAIpQLSef-bIgo0OOr9jGsur9sMggmdq004O8y9Z3Mv64jkY5CCBAUw/viewform?usp=dialog

Vui lòng hoàn thành trước 23:59 ngày 15/07/2025 để đảm bảo tiến độ xét duyệt. Nếu có bất kỳ câu hỏi hoặc cần hỗ trợ, bạn có thể phàn hồi lại với chúng tôi qua email này.

Chúng tôi rất mong được chào đón bạn vào câu lạc bộ và cùng nhau tạo nên những giá trị và trải nghiệm tuyệt vời trong câu lạc bộ!

Trân trọng,  
Ban Tổ chức PAYT

---------------------
Email: paytclub.ptit@gmail.com

"""

ourmailsender = MailSender('paytclub.ptit@gmail.com', mail_key, ('smtp.gmail.com', 587))

ids = set()

for index, row in data.iterrows():
    
    if row['Mã sinh viên'] in ids or not row['Duyệt']:
        continue 
    
    personalized_message = plaintext.format(name=row['Họ và tên'])

    ourmailsender.set_message(personalized_message, "Tiếp Tục Hành Trình Cùng PAYT", "Ban Tổ chức PAYT")
    ourmailsender.set_recipients([row['Email']])
    ourmailsender.connect()
    ourmailsender.send_all()
    
    ids.add(row['Mã sinh viên'])

# ourmailsender.set_message(plaintext.format(name='Nguyễn Minh Quang'), "PAYT - Nộp CV Để Tiếp Tục Hành Trình", "Ban Tổ chức PAYT")
    
# ourmailsender.set_recipients(['nguyenquang71103@gmail.com', 'tran.trung.kien@sun-asterisk.com'])

# ourmailsender.connect()
# ourmailsender.send_all()