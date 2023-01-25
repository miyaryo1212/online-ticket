import ctypes
import smtplib
import tkinter
import tkinter.messagebox
from email import message


def showmsgbox(title, content):
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showinfo(title, content)

    return None


def userinfo():
    root = tkinter.Tk()
    root.withdraw()
    root.title("Enterキーで決定")

    def get(event):
        root.quit()

    label_address = tkinter.Label(root, text="Outlook email address")
    label_address.grid(column=0, row=0, padx=20, pady=10)
    label_password = tkinter.Label(root, text="Password")
    label_password.grid(column=0, row=1, padx=20, pady=10)

    box_address = tkinter.Entry(root, width=30)
    box_address.grid(column=1, row=0, padx=20, pady=10)
    box_password = tkinter.Entry(root, width=30)
    box_password.grid(column=1, row=1, padx=20, pady=10)

    root.bind("<Return>", get)

    root.deiconify()
    root.mainloop()
    root.withdraw()

    if not box_address.get() or not box_password.get():
        showmsgbox("Error", "メールアドレスまたはパスワードが正しく入力されませんでした")

    return box_address.get(), box_password.get()


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass

    username, password = userinfo()
    from_email = username

    smtp_host = "smtp.office365.com"
    smtp_port = 587

    to_email = ""

    msg = message.EmailMessage()
    msg.set_content("Sent from Python")
    msg["Subject"] = "Test"
    msg["From"] = from_email
    msg["To"] = to_email

    server = smtplib.SMTP(smtp_host, smtp_port)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(username, password)
    server.send_message(msg)

    server.quit()
