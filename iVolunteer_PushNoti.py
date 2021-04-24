import smtplib
from email.message import EmailMessage

def email_alert(subject, body, to):
    user = "zippsgame@gmail.com"
    password = "qedqerpxebokpgxh"

    msg = EmailMessage()
    msg.set_content(body)
    msg['subject'] = subject
    msg['to'] = to
    msg['from'] = user
    
    server = smtplib.SMTP("smtp.gmail.com",587)
    server.starttls()
    server.login(user,password)
    server.send_message(msg)
    print("[Event Email Sent!]")
    server.quit()

def eventDetails():
    #pull up Event details -Title -Location -Time -Other Reminders
    return #deets

def notifyBody(emailTo, emailSub, eventTitle):
    eventCurrent = eventTitle
    eventDetails(eventCurrent)
    #emailBod = "Here are the important details for the", eventCurrent[event], "you've signed up for:\n Meet Up Location =", eventCurrent[meetupLoc], "\n Meet Up Time =", eventCurrent[meetupTime], "\n Other Reminder = ", event[meetupReminder]"
    email_alert(emailSub, emailBod, emailTo)

######## Email Test ##########

emailTo = "basilaeadrian@gmail.com" #Volunteer Email
emailSub = "iVolunteer - Event Notification"
emailBod = "[iVolunteer Header & Statement]\n Here are the important details for the [event] you've signed up for:\n\n  Meet Up Location: [meetupLoc]\n  Meet Up Time: [meetupTime]\n  Other Reminders: [meetupReminder] \n\nSee you there! :D\niVolunteer Team"

email_alert(emailSub, emailBod, emailTo) #email w/ qr code
print(emailBod)

######## SMS Test ##########

"""
#sms notification via email carrier gateway
email_alert("Hey", "Server Test", "639052005090@pmms.globe.com.ph") 
"""