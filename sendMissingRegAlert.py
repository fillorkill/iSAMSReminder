import json, datetime, smtplib, requests
from email.message import EmailMessage
from datetime import date

# CONFIG
url = "https://school.isams.cloud/api/batch/1.0/json.ashx?apiKey="
apiKey = "7xxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
# filter only today
start = date.today().strftime("%Y-%m-%d")
end = (date.today() + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
# RegistrationDateTimes to send email notifications for
prePrepAM = start + 'T08:40:00'
prepAM = start + 'T08:00:00'
seniorAM = start + 'T08:00:00'
boardingPM = start + 'T17:30:00'
registrationDateTimesToNotify = [prePrepAM, prepAM, seniorAM, boardingPM]
boardingNotificationEmail = 'boarding@school.com.cn'
additionalEmailTo = 'additionalEmails@school.com.cn'
### send email alert
sender = 'noreply@school.com.cn'
password = 'password'
host = 'smtp.office365.com'
port = 587
secureConnection = False
# path
response_path = '/opt/isams/response.json'
output_path = '/opt/isams/missingRegAlertOutput.json'

def getResponse(jsonStorePath, apiKey, filter=""):
    try:
        # get isams json
        print("Connecting to iSAMS...")
        if filter is None or filter == "":
            response = requests.get(url + apiKey)
        else:
            headers = {'Content-Type': 'application/xml'}
            response = requests.post(url + apiKey, data=filter, headers=headers)
        with open(jsonStorePath, "w", encoding="utf8") as f:
            json.dump(response.json(), f, ensure_ascii=False)
    except:
        print("Error connecting to iSAMS API. Abort.")
        response = None
    return response

filter = """<?xml version="1.0" encoding="utf-8" ?>
<Filters>
         <Registration>
                 <RegistrationStatus StartDate='""" + start + """' EndDate='""" + end + """' />
         </Registration>
</Filters>"""

r = getResponse(response_path, apiKey, filter).json()

unregistered = {}
unregistered_studentAffairs = {}

if 'iSAMS' in r and \
   r['iSAMS'] is not None and \
   'RegistrationManager' in r['iSAMS'] and \
   r['iSAMS']['RegistrationManager'] is not None and \
   'RegistrationStatuses' in r['iSAMS']['RegistrationManager'] and \
   r['iSAMS']['RegistrationManager']['RegistrationStatuses'] is not None and \
   'RegistrationStatus' in r['iSAMS']['RegistrationManager']['RegistrationStatuses'] and \
   r['iSAMS']['RegistrationManager']['RegistrationStatuses']['RegistrationStatus'] is not None:
    for registration in r['iSAMS']['RegistrationManager']['RegistrationStatuses']['RegistrationStatus']:
        #get unregistered
        if registration['Registered'] == "0" and 'Code' not in registration:
            if registration['RegistrationDateTime'] in registrationDateTimesToNotify:
                # get student details
                for pupil in r['iSAMS']['PupilManager']['CurrentPupils']['Pupil']:
                    if pupil['SchoolId'] == registration['PupilId']:
                        # boarding, do no action if not past registration time
                        if registration['RegistrationDateTime'] == boardingPM and datetime.datetime.strptime(boardingPM, "%Y-%m-%dT%H:%M:%S") < datetime.datetime.now():
                            # for student affairs
                            unregistered_studentAffairs.update({"Boarding - " + registration['@Id']: registration['RegistrationDateTime'] + " ::: " + pupil['Form'] + " - " + 
                                pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                ") - Boarding - "})
                            # for indivdual teachers
                            if boardingNotificationEmail not in unregistered:
                                unregistered[boardingNotificationEmail] = {}
                            unregistered[boardingNotificationEmail].update({registration['@Id']: registration['RegistrationDateTime'] + " ::: " + pupil['Form'] + " - " + 
                                pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                ") - Boarding - "})
                        # non-boarding pre-prep and prep, by form
                        elif (pupil['DivisionName'] == "Prep" or pupil['DivisionName'] == "Pre-Prep") and registration['RegistrationDateTime'] != boardingPM:
                            emails = []
                            # get form tutor
                            for form in r['iSAMS']['SchoolManager']['Forms']['Form']:
                                if form['Form'] == pupil['Form']:
                                    # get form tutor email address
                                    tutor = ""
                                    asstTutor = ""
                                    secondAsstTutor = ""
                                    if '#text' in form['Tutor']:
                                        tutor = form['Tutor']['#text']
                                    if '#text' in form['AssistantFormTutor']:
                                        asstTutor = form['AssistantFormTutor']['#text']
                                    if '#text' in form['SecondAssistantFormTutor']:
                                        secondAsstTutor = form['SecondAssistantFormTutor']['#text']
                                    for staff in r['iSAMS']['HRManager']['CurrentStaff']['StaffMember']:
                                        if staff['UserCode'] in [tutor, asstTutor, secondAsstTutor]:
                                            emails.append(staff['SchoolEmailAddress'])
                                    break
                            # for student affairs
                            unregistered_studentAffairs.update({pupil['Form'] +" - " + registration['@Id']: registration['RegistrationDateTime'] + " ::: " + pupil['Form'] + " - " + 
                                    pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                    ")"})
                            
                            # for individual teachers
                            for email in emails:
                                if email not in unregistered:
                                    unregistered[email] = {}
                                unregistered[email].update({registration['@Id']: registration['RegistrationDateTime'] + " ::: " + pupil['Form'] + " - " + 
                                    pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                    ")"})
                            
                        # non-boarding senior, by tutor group
                        elif pupil['DivisionName'] == "Senior" and registration['RegistrationDateTime'] != boardingPM:
                            for staff in r['iSAMS']['HRManager']['CurrentStaff']['StaffMember']:
                                if staff['UserCode'] == pupil['Tutor']:
                                    # for student affairs
                                    unregistered_studentAffairs.update({staff['FullName'] + " - " + registration['@Id']: registration['RegistrationDateTime'] + " ::: " + staff['FullName'] + " - " + 
                                    pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                    ")"})

                                    # for individual teachers
                                    if staff['SchoolEmailAddress'] not in unregistered:
                                        unregistered[staff['SchoolEmailAddress']] = {}
                                    unregistered[staff['SchoolEmailAddress']].update({registration['@Id']: registration['RegistrationDateTime'] + " ::: " + staff['FullName'] + " - " + 
                                    pupil['Surname'] + ", " + pupil['Forename'] + " (" + pupil['Preferredname'] + 
                                    ")"})
                                    break
                        break

with open(output_path, "w") as f:
    json.dump(unregistered, f)

with smtplib.SMTP(host, port) as smtp:
        smtp.starttls()
        smtp.login(sender, password)
        
        # full copy sent to student services
        if additionalEmailTo != "":
            msg = EmailMessage()
            msg['Subject'] = "Missing Registration Alert - " + start
            msg['From'] = sender
            msg['To'] = 'student.affairs@kings-school.com.cn'
            msg.set_content("STUDENT AFFAIRS COPY\r\nThis email is sent automatically to all form tutors in your class/group twice a day and contains missing registration data in the morning and boarding in the afternoon.\r\n" 
            +"A full copy is sent to student affairs for record keeping.\r\n"
            + "The date and times below represent the missing registration period. This is typically the first registration in the morning and for boarding, it is the boarding registration period. \r\n"
            +"The following students have missing registration on " + start + " as of "+datetime.datetime.now().strftime("%H:%M")+":\r\n" 
            + json.dumps(unregistered_studentAffairs, indent=4, sort_keys=True) + "\r\n"
            +"If you have registered these students today, please ignore this email.")
            smtp.send_message(msg)
            print("Alert sent to " + additionalEmailTo + ".")

        # sent to individual teachers
        for email in unregistered:
            msg = EmailMessage()
            msg['Subject'] = "Missing Registration Alert - " + start
            msg['From'] = sender
            msg['To'] = email
            msg.set_content("This email is sent automatically to all form tutors in your class/group twice a day and contains missing registration data in the morning and boarding in the afternoon.\r\n"+
                            "A full copy is sent to student affairs for record keeping.\r\n" + 
                            "The date and times below represent the missing registration period. This is typically the first registration in the morning and for boarding, it is the boarding registration period. \r\n" +
                            "The following students have missing registration on " + start + " as of "+datetime.datetime.now().strftime("%H:%M")+":\r\n" 
                                       + json.dumps(unregistered[email], indent=4) + "\r\n" +
                                       "If you have registered these students today, please ignore this email.")
            smtp.send_message(msg)
            print("Alert sent to " + email + '.')
        smtp.quit()