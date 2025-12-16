#!/usr/bin/env python
# coding: utf-8

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os

def send_email(subject, message, recipient, attachment_path):
    try:
        sender_email = 'youremai;@email.com'
        sender_password = 'google-generated-password'
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient
        msg['Subject'] = subject

        msg.attach(MIMEText(message, 'html'))

        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(attachment_path)}')
            msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()

    except Exception as e:
        print(f"Failed to send email to {recipient}. Error: {e}")

def load_data_from_excel(file_path):
    try:
        df_scores = pd.read_excel(file_path, sheet_name=2)
        return df_scores
    except Exception as e:
        print(f"Error loading Excel data: {e}")
        return None

def main():
    ######################################
    file_path = 'some_file_path'
    attachment_path = 'your_attachment'
    ######################################

    if not os.path.exists(attachment_path):
        print(f"Attachment file '{attachment_path}' not found.")
        return

    data = load_data_from_excel(file_path)

    if data is not None:
        for index, student in data.iterrows():
            personalized_info = f"\n\nชื่อผู้เข้าสอบ: {student['name']} {student['surname']}\nโรงเรียน: {student['school']}"
            html_message = f"""
                <p>เรียน ผู้เข้าร่วมการแข่งขัน <strong>MU Mental Math Competition 2025</strong> ทุกท่าน</p>

                <p>ในอีเมลนี้ ประกอบไปด้วย <strong>คำถามที่พบบ่อย / ตารางการแข่งขัน / เล่มกติกา / สถานที่จัดแข่ง</strong></p>

                <p>การแข่งขันจะจัดขึ้นในวันพรุ่งนี้ (วันเสาร์ที่ 19 กรกฎาคม 2568) ณ อาคารบรรยายรวม คณะวิทยาศาสตร์ มหาวิทยาลัยมหิดล วิทยาเขตพญาไท</p>

                <ul>
                    <li>แผนที่ Google Maps: <a href="https://maps.app.goo.gl/sT9kwdaWBbPjRYae7" target="_blank">คลิกที่นี่</a></li>
                    <li>แผนที่ภายในมหาวิทยาลัย: <a href="https://science.mahidol.ac.th/th/map.php" target="_blank">คลิกที่นี่</a></li>
                    <li>ข้อมูลเพิ่มเติมและคำถามที่พบบ่อย (ที่จอดรถ, การแต่งกาย ฯลฯ): <a href="https://www.facebook.com/share/p/15iraC6UuA/" target="_blank">คลิกที่นี่</a></li>
                </ul>

                <p>สามารถตรวจสอบ <strong>ตารางเวลาการแข่งขัน</strong> และ <strong>เล่มกติกา</strong> ได้ที่ไฟล์แนบท้ายอีเมลนี้ (ตารางแข่งขันอยู่หน้าสุดท้ายของเล่มกติกา)</p>

                <p><strong>คะแนนการแข่งขัน</strong> จะประกาศผ่านทาง Facebook: <strong>MATH MU</strong> และ Instagram: <strong>mentalmath_2025</strong> รวมถึงแจ้งคะแนนรายบุคคลทางอีเมล</p>

                <p>หลังประกาศคะแนนแล้ว หากต้องการขอดูกระดาษคำตอบ จะมีเวลา 24 ชั่วโมงในการติดต่อผ่าน Facebook หรือ Instagram</p>

                <p>ติดตามข่าวสารล่าสุดได้จากเพจ: 
                    <ul>
                        <li>Facebook: <strong><a href="https://www.facebook.com/MATHMAHIDOLU" target="_blank">MATH MU</a></strong></li>
                        <li>Instagram: <strong><a href="https://www.instagram.com/mentalmath_2025/" target="_blank">mentalmath_2025</a></strong></li>
                    </ul>
                </p>

                <br><br>อีเมลนี้ถูกสร้างขึ้นโดยอัตโนมัติ หากมีข้อสงสัยเพิ่มเติม สามารถสอบถามนักศึกษาเจ้าของโครงการ MU Mental Math Competition (นายอภิวัฒน์ คงสวัสดิ์)<br>
                ได้ที่อีเมล:

                <br><br>--
                <br><br><strong>Apiwat Kongsawat</strong>, Founder of MU Mental Math Competition<br>
                Department of Mathematics, Mahidol University<br>
                Rama VI Road, Ratchathewi, Bangkok 10400, Thailand<br>
                """

            subject = "แจ้งข้อมูลสำคัญของงานแข่ง MU Mental Math Competition 2025"
    
            send_email(subject, html_message, student['email'], attachment_path)
            print(f"Email successfully sent to {student['email']} (Count = {index+1})")

if __name__ == "__main__":
    main()
