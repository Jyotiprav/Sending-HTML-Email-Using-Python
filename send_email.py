#Inspired By :https://gist.github.com/perfecto25/4b79b960eb58dc1f6025b56394b51cc1

# way of using operating system based functionality
import os
# Way of getting system information
import sys
# jinja2 is mark_up language for Python developers. The following function will render the html template into jinja2.
def load_html_template_as_jinja(template, **kwargs):
    # check if file exists
    if not os.path.exists(template):
        print('No template file present: %s' % template)
        sys.exit()
    #if file exits change it to jinja2
    import jinja2
    jinja_loader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=jinja_loader)
    final = templateEnv.get_template(template)
    return final.render(**kwargs)


# ------------------------------------------------------------------------------------------------
def send_email(to, sender='MyCompanyName<noreply@mycompany.com>', cc=None, bcc=None, subject=None, body=None):
    ''' sends email using a Jinja HTML template '''
    import smtplib
    # Import the email modules
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    # convert TO into list if string
    if type(to) is not list:
        to = to.split()

    to_list = to + [cc] + [bcc]
    to_list = filter(None, to_list)  # remove null emails

    msg = MIMEMultipart('alternative')
    msg['From'] = sender
    msg['Subject'] = subject
    msg['To'] = ','.join(to)
    msg['Cc'] = cc
    msg['Bcc'] = bcc
    msg.attach(MIMEText(body, 'html'))
    server = smtplib.SMTP('smtp.gmail.com:587')  # or your smtp server
    # Print debugging output when testing

    server.set_debuglevel(1)

    # Credentials (if needed) for sending the mail
    password = "put_your_password_here"

    server.starttls()
    server.login(sender, password)
    server.sendmail(sender, to, msg.as_string())
    server.quit()

# MAIN
import xlrd
to_list = []
loc = ("project1/email_id-sheet.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
for i in range(sheet.nrows):
    print(sheet.cell_value(i, 0))
    to_list.append(sheet.cell_value(i, 0))
# generate HTML from template
html = load_html_template_as_jinja('project1/scalar.html', vars=locals())
sender = 'sharma.jyoti91@gmail.com'
cc = None
bcc = None
subject = 'Scalars Academy'
# send email to a list of email addresses
send_email(to_list, sender, cc, bcc, subject, html)