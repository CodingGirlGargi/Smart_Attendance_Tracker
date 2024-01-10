import face_recognition
import cv2
import numpy as np
import os
import xlwt
from xlwt import Workbook
from datetime import datetime
import xlrd, xlwt
from xlutils.copy import copy as xl_copy
import dlib
import openpyxl

def get_empty_columns(filename, sheetname):
    # Load the Excel workbook
    wb = openpyxl.load_workbook(filename)
    
    # Select the sheet by name
    sheet = wb[sheetname]

    empty_columns = []

    # Iterate through columns in the selected sheet
    for column in sheet.iter_cols(min_col=1, max_col=sheet.max_column):
        # Check if all cells in the column are empty
        if all(cell.value is None for cell in column):
            empty_columns.append(column[0].column)

    return empty_columns
def search_and_write_value(sheet, search_value,column_Number):
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            # Check if the cell contains the search value
            if cell.value == search_value:
                # Write 1 to the column beside it
                sheet.cell(row=cell.row, column=column_Number, value=1)

    # Save the changes to the Excel file
    wb.save(fileName)


CurrentFolder = os.getcwd()

image1 = CurrentFolder+'\\SudhaMurty.jpg'
image2 = CurrentFolder+'\\RatanTata.jpg'
image3 = CurrentFolder+ '\\NarayanaMurthy.jpg'


# This is a demo of running face recognition on live video from your webcam. It's a little more complicated than the
# other example, but it includes some basic performance tweaks to make things run a lot faster:
#   1. Process each video frame at 1/4 resolution (though still display it at full resolution)
#   2. Only detect faces in every other frame of video.

# PLEASE NOTE: This example requires OpenCV (the `cv2` library) to be installed only to read from your webcam.
# OpenCV is not required to use the face_recognition library. It's only required if you want to run this
# specific demo. If you have trouble installing it, try any of the other demos that don't require it instead.

# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)

# Load a sample picture and learn how to recognize it.
person1_name = "SudhaMurty"
person1_image = face_recognition.load_image_file(image1)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

# Load a second sample picture and learn how to recognize it.
person2_name = "RatanTata"
person2_image = face_recognition.load_image_file(image2)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

person3_name = "NarayanaMurthy"
person3_image = face_recognition.load_image_file(image3)
person3_face_encoding = face_recognition.face_encodings(person3_image)[0]


# Create arrays of known face encodings and their names
known_face_encodings = [
    person1_face_encoding,
    person2_face_encoding,person3_face_encoding]
known_face_names = [
    person1_name,
    person2_name,person3_name]

# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

'''
rb = xlrd.open_workbook('attendence_excel.xls', formatting_info=True) 
wb = xl_copy(rb)
inp = input('Please give current subject lecture name')
sheet1 = wb.add_sheet(inp)
sheet1.write(0, 0, 'Name/Date')
sheet1.write(0, 1, str(date.today()))
row=1
col=0
already_attendence_taken = ""
already_present = []
'''
already_attendence_taken = ""
fileName = 'Teacher.xlsx'
sheetName = input("Please enter the subject and class: ")

wb = openpyxl.load_workbook(fileName)
sheet = wb[sheetName]
empty_columns = get_empty_columns(fileName,sheetName)
column_Number = empty_columns[0]

current_date = datetime.now().date()
sheet.cell(row = 1, column = empty_columns[0], value = current_date)


while True:
            # Grab a single frame of video
            ret, frame = video_capture.read()

            # Resize frame of video to 1/4 size for faster face recognition processing
            small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)

            # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
            rgb_small_frame = small_frame[:, :, ::-1]

            # Only process every other frame of video to save time
            if process_this_frame:
                # Find all the faces and face encodings in the current frame of video
                face_locations = face_recognition.face_locations(rgb_small_frame)
                face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations, model="small")

                face_names = []
                for face_encoding in face_encodings:
                    # See if the face is a match for the known face(s)
                    matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                    name = "Unknown"

                    # # If a match was found in known_face_encodings, just use the first one.
                    # if True in matches:
                    #     first_match_index = matches.index(True)
                    #     name = known_face_names[first_match_index]

                    # Or instead, use the known face with the smallest distance to the new face
                    face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                    best_match_index = np.argmin(face_distances)
                    if matches[best_match_index]:
                        name = known_face_names[best_match_index]

                    face_names.append(name)
                    if((already_attendence_taken != name) and (name != "Unknown")):
                        search_and_write_value(sheet,name,column_Number)
                        print("attendence taken")
                        already_attendence_taken = name                       
   
                    else:
                     print("next student")
                        
            process_this_frame = not process_this_frame


            # Display the results
            for (top, right, bottom, left), name in zip(face_locations, face_names):
                # Scale back up face locations since the frame we detected in was scaled to 1/4 size
                top *= 4
                right *= 4
                bottom *= 4
                left *= 4

                # Draw a box around the face
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

                # Draw a label with a name below the face
                cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
                font = cv2.FONT_HERSHEY_DUPLEX
                cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

            # Display the resulting image
            cv2.imshow('Video', frame)

            # Hit 'q' on the keyboard to quit!
            if cv2.waitKey(1) & 0xff==ord('q'):   
                print("data save")
                break

# Release handle to the webcam
video_capture.release()     
cv2.destroyAllWindows()

'''
                     sheet1.write(row, col, name )
                     col =col+1
                     sheet1.write(row, col, "Present" )
                     row = row+1
                     col = 0
                     print("attendence taken")
                     wb.save('attendence_excel.xls')
                     already_attendence_taken = name
                     already_present.append(name)'''