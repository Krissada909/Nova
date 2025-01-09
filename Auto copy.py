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
            print(f"Marking as read: {message.Subject}")
            
            # ทำเครื่องหมายว่าอ่านแล้ว (แทนการใช้ MarkAsRead)
            message.Unread = False  # ทำเครื่องหมายว่าอ่านแล้ว
            message.Save()  # บันทึกการเปลี่ยนแปลง
            
            # สร้างอีเมลตอบกลับ
            print(f"Replying to: {message.Subject}")
            reply = message.Reply()  # ใช้ Reply แทน ReplyAll
            reply.Body = reply_message + "\n\n" + reply.Body  # เพิ่มข้อความในอีเมล
            reply.Send()  # ส่งอีเมล
            print("Email replied successfully.")

# เรียกใช้งาน
folder_name = "Test"  # ชื่อโฟลเดอร์ที่ต้องการ
reply_message = "ขอบคุณสำหรับอีเมลของคุณ เราจะติดต่อกลับโดยเร็วที่สุด!"
auto_reply_outlook(folder_name, reply_message)
