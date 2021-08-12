import os
import smtplib
import openpyxl as xl
from email.message import EmailMessage
# username = str(input('Your Username:' ))
# password = str(input('Your Password:' ))
EMAIL_ADDRESS = "******@gmail.com	"
EMAIL_PASSWORD = "*******"
wb = xl.load_workbook('list.xlsx')
sheet1 = wb.get_sheet_by_name('Sheet1')
print("Starting anyway.....")
names = []
emails = []
interests=[]
des=[]
inst = []
for cell in sheet1['B']:
    emails.append(cell.value)

for cell in sheet1['D']:
    inst.append(cell.value)
for cell in sheet1['C']:
    names.append(cell.value)
for cell in sheet1['A']:
    des.append(cell.value)

print("Data collection done!")

# print(len(emails))
# print(len(inst))
# print(len(names))
# print(len(interests))
# print(len(particular))

for i in range(len(emails)):
	msg = EmailMessage()
	# print(i)
	# msg.set_content('hello')
	
	msg.add_alternative("""<!DOCTYPE html>
<html>
    <body>
<p>Respected <strong>"""+des[i]+""" """+names[i]+""",</strong></p>
<p>MAil Content</p>

<p>Yours Sincerely,<br>
XYZ<br>
Sophomore ,XYZ</p>



    </body>
</html>""", subtype='html')
	msg['Subject'] = 'Summer Intern'
	msg['From'] = EMAIL_ADDRESS
	msg['To'] = emails[i]
	print("Message Proccessed!")
	print("Sending to "+emails[i])
	with open('Resume.pdf','rb') as res:
		msg.add_attachment(res.read(),maintype='application',subtype='pdf',filename='Resume.pdf')
	with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
		smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
		smtp.send_message(msg)
		# print(msg)
		print('Mail sent to', emails[i])
print('All emails sent successfully!')

    


