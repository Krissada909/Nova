import win32com.client
import re

def auto_reply_outlook(folder_name):
    try:
        # เริ่มต้นการเชื่อมต่อ Outlook
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # เลือกบัญชีอีเมลที่ต้องการ
        print("Fetching account...")
        account = namespace.Folders.Item("krissada.s@ztrus.com")
        print(f"Selected account: {account.Name}")
        
        # หาโฟลเดอร์ย่อยตามชื่อ
        folder = None
        for subfolder in account.Folders:
            print(f"Checking folder: {subfolder.Name}")
            if subfolder.Name.lower() == folder_name.lower():
                folder = subfolder
                break
        
        if folder is None:
            print(f"Folder '{folder_name}' not found!")
            return
        
        print(f"Found folder: {folder.Name}")
        
        # ดึงอีเมลในโฟลเดอร์
        messages = folder.Items
        print(f"Number of messages: {len(messages)}")
        
        # ใช้การเรียงลำดับใหม่ให้ดึงอีเมลเก่าที่สุดก่อน
        messages.Sort("[ReceivedTime]", False)
        
        for message in messages:
            if message.Unread:  # ตรวจสอบว่าอีเมลยังไม่ได้อ่าน
                print(f"Processing unread email: {message.Subject}")
                
                # ดึงหมายเลข INC จากหัวข้ออีเมล
                match = re.search(r'INC-\d+', message.Subject)
                ticket_type = match.group() if match else "N/A"
                
                # ตรวจสอบคำสำคัญในหัวข้ออีเมล
                keywords = [
                    "อ่านสินค้าไม่ตรงกับ PO",
                    "ราคาต่อหน่วยอ่านไม่ถูกต้อง",
                    "บรรทัดสุดท้ายอ่านไม่ถูกต้อง",
                    "อ่านสินค้าผิด",
                    "อ่านราคาต่อหน่วยไม่ตรงตามเอกสาร",
                    "อ่าน Desc เกิน",
                    "อ่านไม่ครบ",
                    "อ่านขาด",
                    "อ่านไม่ตรง",
                    "อ่านผิด",
                    "อ่านข้อมูลไม่ถูกต้อง",
                    "อ่านข้อมูลผิด",
                    "อ่านจำนวนสินค้าผิด"
                ]
                urgency = "Medium"
                if any(keyword in message.Subject for keyword in keywords):
                    urgency = "Low"
                    ticket_type = f"SR-{ticket_type.split('-')[1]}" if ticket_type != "N/A" else "SR-N/A"
                
                print(f"Urgency: {urgency}, Ticket Type: {ticket_type}")
                
                # สร้างอีเมลตอบกลับ
                reply = message.Reply()
                reply.To = "nova@dksh.com; kunchalee.s@dksh.com; kusak.k@dksh.com"  # ส่งเฉพาะไปยังผู้รับนี้
                reply.CC = "narakorn.p@n2nsp.com; sumate.s@n2nsp.com; ithipan.m@ztrus.com; sirima.s@ztrus.com; doungkamol.j@ztrus.com; panachit.k@ztrus.com"  # เพิ่มผู้รับในช่อง Cc
                
                
                # เพิ่มข้อความตอบกลับที่มีตาราง HTML ไปที่ด้านบนของเนื้อหาอีเมล
                reply.HTMLBody = f"""
                <p>Dear {message.SenderName},</p>
                <p>Thanks for reaching out! Your request ({ticket_type}) has been received and is being reviewed by our support staff.</p>
                <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">
                    <tr>
                        <th>Source</th>
                        <th>Urgency</th>
                        <th>Ticket Type</th>
                        <th>Resolution</th>
                        <th style="width: 40%; text-align: left;">Tags</th>
                    </tr>
                    <tr>
                        <td>OCR</td>
                        <td>{urgency}</td>
                        <td>{ticket_type}</td>
                        <td></td>
                        <td style="width: 40%; text-align: left;"></td>
                    </tr>
                </table>
                """ + reply.HTMLBody
                
                # ส่งอีเมล
                  # ระบุบัญชีที่จะใช้ส่ง
                for acc in outlook.Session.Accounts:
                    if acc.SmtpAddress.lower() == "krissada.s@ztrus.com":
                        reply._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                        break
                reply.Send()

                print("Email replied successfully.")
                
                # ทำเครื่องหมายว่าอ่านแล้ว
                message.Unread = False
                message.Save()  # บันทึกการเปลี่ยนแปลง
    except Exception as e:
        print(f"An error occurred: {e}")

# เรียกใช้งาน
folder_name = "Test"  # ชื่อโฟลเดอร์ที่ต้องการ
auto_reply_outlook(folder_name)
