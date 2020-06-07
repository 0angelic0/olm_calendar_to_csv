import os
import csv
import xml.etree.ElementTree as ET

appointments = []
filename = "Calendar.xml"
fieldnames = ['Summary', 'Organizer', 'StartTime', 'EndTime', 'Description', 'Attendees']

try:
  tree = ET.parse(filename)
except:
  print(filename+" couldn't parsed")
finally:
  root = tree.getroot()
  for appointment in root.iter('appointment'):
    to_append = {}
    if appointment.find('OPFCalendarEventCopySummary') is not None:
      to_append['Summary'] = appointment.find('OPFCalendarEventCopySummary').text
    if appointment.find('OPFCalendarEventCopyOrganizer') is not None:
      to_append['Organizer'] = appointment.find('OPFCalendarEventCopyOrganizer').text
    if appointment.find('OPFCalendarEventCopyStartTime') is not None:
      to_append['StartTime'] = appointment.find('OPFCalendarEventCopyStartTime').text
    if appointment.find('OPFCalendarEventCopyEndTime') is not None:
      to_append['EndTime'] = appointment.find('OPFCalendarEventCopyEndTime').text
    # if appointment.find('OPFCalendarEventCopyDescriptionPlain') is not None:
    #   to_append['Description'] = appointment.find('OPFCalendarEventCopyDescriptionPlain').text
    to_append['Attendees'] = []
    attendeeList = appointment.find('OPFCalendarEventCopyAttendeeList')
    if attendeeList is not None:
      for attendee in attendeeList.findall('appointmentAttendee'):
        to_append['Attendees'].append(attendee.attrib['OPFCalendarAttendeeAddress'])
    appointments.append(to_append)

w = csv.DictWriter(open("output.csv", "w"), fieldnames=fieldnames, dialect='excel-tab')
w.writeheader()
w.writerows(appointments)
print("operation completed")