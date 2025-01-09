import win32com.client
import re

def auto_reply_outlook(folder_name):
    try:
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("Fetching account...")
        account = namespace.Folders.Item("krissada.s@ztrus.com")
        print(f"Selected account: {account.Name}")
        
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
        messages = folder.Items
        print(f"Number of messages: {len(messages)}")
        
        messages.Sort("[ReceivedTime]", False)
        success_count = 0

        for message in messages:
            if message.Unread:
                print(f"Processing unread email: {message.Subject}")
                
                match = re.search(r'INC-\d+', message.Subject)
                ticket_type = match.group() if match else "N/A"
                
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
                    "อ่านจำนวนสินค้าผิด",
                    "อ่านจำนวนและราคาผิด",
                    "PO ไม่ถูกต้อง",
                    "อ่านรายการแถมไม่ตรง",
                    "Usage ไม่ถูกต้องตามเอกสาร",
                    "อ่านเลข PO ไม่ถูกต้อง",
                    "อ่านจำนวนขายรวมแถม และราคาไม่ตรง",
                    "อ่านไม่ถูกต้องตามเอกสาร",
                    "อ่าน SO Form No ไม่ตรง",
                    "QA Qty ไม่ตรงตามเอกสาร",
                    "อ่านสินค้าคนละรายการกับเอกสารระบุ",
                    "โค้ดสินค้าไม่ตรงกับสินค้าใน PO",
                    "อ่านข้อความที่ไม่เกี่ยวข้องมาด้วย",
                    "อ่านสินค้าคนละรายการกับเอกสารระบุ",
                    "รายการที่ 1 อ่านสินค้าไม่ตรง",
                    "อ่าน Desc เกิน","อ่านเลขที่ PO ไม่ครบ",
                    "ราคาต่อหน่วยไม่ถูกต้องตามเอกสาร",
                    "อ่านเลขที่ PO ผิด, อ่าน Desc ผิด"
                ]
                
                urgency = "Medium"
                found_keyword = False
                for keyword in keywords:
                    if keyword in message.Subject:
                        urgency = "Low"
                        found_keyword = True
                        ticket_type = f"SR-{ticket_type.split('-')[1]}" if ticket_type != "N/A" else "SR-N/A"
                        break
                
                if found_keyword:
                    parts = message.Subject.split('/')
                    extra_keywords = ["Qty แถมเป็นขาย", "เอกสารไม่มีระบุ Payment","อ่าน C.Product Code มารวมกับ Desc","ไม่อ่าน UOM","UOM รายการที่ 2 ไม่อ่าน","ไม่อ่านรายการแถม",
                    "อ่าน UOM ไม่ครบทุกรายการ","Payment ไม่ได้เลือก แต่ขึ้น ZPA0","อ่าน C.Product Code มารวมกับ Desc","อ่านสินค้าไม่ครบ มี 2 อ่าน 1 รายการ","อ่าน UOM ไม่ครบทุกรายการ",
                    "หย่อน UOM ผิดช่อง"]
                    for part in parts[1:]:
                        if any(extra_keyword in part for extra_keyword in extra_keywords):
                            urgency = "Medium"
                            ticket_type = f"INC-{ticket_type.split('-')[1]}" 
                            break
                
                print(f"Urgency: {urgency}, Ticket Type: {ticket_type}")
                
                reply = message.Reply()
                reply.To = "nova@dksh.com; kunchalee.s@dksh.com; kusak.k@dksh.com"  # ส่งเฉพาะไปยังผู้รับนี้
                reply.CC = "narakorn.p@n2nsp.com; sumate.s@n2nsp.com; ithipan.m@ztrus.com; sirima.s@ztrus.com; doungkamol.j@ztrus.com; panachit.k@ztrus.com"  # เพิ่มผู้รับในช่อง Cc

           
               
                
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
                <br><br>
                <p>Best Regards,<br>
                Krissada Sangkampai<br><br>
                ZTRUS SUPPORT TEAM</p>
                <p><a href='http://www.n2nsp.com'>www.n2nsp.com</a> | +66 815570105
                """ + reply.HTMLBody
                
                for acc in outlook.Session.Accounts:
                    if acc.SmtpAddress.lower() == "krissada.s@ztrus.com":
                        reply._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                        break
                reply.Send()
                print("Email replied successfully.")
                
                message.Unread = False
                message.Save()

                success_count += 1
            
        print(f"Total emails replied successfully: {success_count}")
    except Exception as e:
        print(f"An error occurred: {e}")

class Export_NOVA_Ticket_AI_OCR():
    def dattime ():
        print("NOVA Ticket AI_OCR 08.01.2025 11:25")
        pass
        

    pass
    


folder_name = "Test"
auto_reply_outlook(folder_name)
