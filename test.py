import win32com.client

def auto_reply_outlook(folder_name, reply_message):
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
            print(f"Replying to unread email: {message.Subject}")
            
            # สร้างอีเมลตอบกลับ
            reply = message.Reply()
            
            # เพิ่มข้อความตอบกลับไปที่ด้านบนของเนื้อหาอีเมล
            reply.HTMLBody = f"<p>{reply_message}</p>" + reply.HTMLBody  # ใช้ HTMLBody เพื่อคงรูปแบบต้นฉบับไว้
            
            # ดึงซิกเนเจอร์ของผู้ใช้งาน
            signature = outlook.Session.CurrentUser.AutoSignature
            
            # เพิ่มซิกเนเจอร์ที่ด้านล่างของอีเมล
            reply.HTMLBody += f"<p>{signature}</p>"
            
            # ส่งอีเมล
            reply.Send()
            print("Email replied successfully.")
            
            # ทำเครื่องหมายว่าอ่านแล้ว
            message.Unread = False
            message.Save()  # บันทึกการเปลี่ยนแปลง

# เรียกใช้งาน
folder_name = "Test"  # ชื่อโฟลเดอร์ที่ต้องการ
reply_message = "ขอบคุณสำหรับอีเมลของคุณ เราจะติดต่อกลับโดยเร็วที่สุด!"
auto_reply_outlook(folder_name, reply_message)