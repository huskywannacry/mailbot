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

# Nội dung email
plaintext = """Xin chào {name},

Cảm ơn bạn đã dành thời gian tham gia vòng sơ loại của Câu lạc bộ Tài năng trẻ AI PTIT (PAYT). Chúng tôi rất trân trọng sự quan tâm và nỗ lực của bạn trong quá trình tuyển chọn.

Sau khi xem xét kỹ lưỡng, rất tiếc chúng tôi phải thông báo rằng hồ sơ của bạn chưa phù hợp với các tiêu chí để tiến vào vòng tiếp theo của đợt tuyển chọn lần này. Tuy nhiên, điều này không làm giảm giá trị của những kỹ năng và đam mê mà bạn đã thể hiện.

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

for index, row in data.iterrows():
    
    if row['Mã sinh viên'] in ids or row['Duyệt']:
        continue 
    
    personalized_message = plaintext.format(name=row['Họ và tên'])

    ourmailsender.set_message(personalized_message, "Thông báo kết quả vòng hồ sơ", "Ban Tổ chức PAYT")
    ourmailsender.set_recipients([row['Email']])
    ourmailsender.connect()
    ourmailsender.send_all()
    
    ids.add(row['Mã sinh viên'])

# ourmailsender.set_message(plaintext.format(name='Nguyễn Minh Quang'), "Thông báo kết quả vòng hồ sơ", "Ban Tổ chức PAYT")
    
# ourmailsender.set_recipients(['nguyenquang71103@gmail.com', 'tran.trung.kien@sun-asterisk.com'])

# ourmailsender.connect()
# ourmailsender.send_all()