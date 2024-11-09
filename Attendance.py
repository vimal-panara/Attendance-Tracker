from asyncore import write
import pytesseract as tess
import cv2
import pandas as pd
import re
from datetime import datetime as dt


def find_names(names_image):
	# function for finding names in proper format from list.

	# Regular Expression patter for extracting names from collected data
	pattern = r'([\w]+\s[\w]+).*'

	final_names = []
	# Iterating over all entries found with the help of cv2
	for i in range(len(names_image)):
		# eliminating entries which are empty
		if len(names_image[i]) > 0:
			# searching for name and surname with the help of regex
			name_found = re.search(pattern, names_image[i])
			final_names.append(name_found[1])

	return final_names


def merge_two_list(list_1, list_2):
	# Function to merge two list of students from two different screenshots

	# we can not just add list_1 to final_list as some student may join from two devices
	final_list = []

	# iterating over first list so that we can remove duplicates
	for item in list_1:
		if item not in final_list:
			final_list.append(item)
	# iterating over second list so that we can remove duplicates
	for item in list_2:
		if item not in final_list:
			final_list.append(item)

	return final_list


def mark_attendance(students_present):
	# Function to mark attendace in the excel sheet

	# date = (str(dt.now()).split(' ')[0]).split('-')
	# date = f'{date[2]}.{date[1]}.{date[0][2:]}'

	# writer = pd.ExcelWriter(r'D:\Study\IITGN\Sem 4\student_list.xlsx')
	# students_present = pd.DataFrame(students_present)
	# students_present.to_excel(writer, sheet_name='Attendance', header=['Names'], index=False)
	# writer.save()

	df = pd.read_excel(r'D:\Study\IITGN\Sem 4\student_list.xlsx', sheet_name='Attendance')
	names = df['Names']
	attendace_status = []
	for student in names:
		if student in students_present:
			attendace_status.append(1)
		else:
			attendace_status.append(0)
	df2 = pd.DataFrame(attendace_status, columns=[date])
	df = df.join(df2)
	df.to_excel(r'D:\Study\IITGN\Sem 4\student_list.xlsx', sheet_name='Attendance', index=False)


# locating the tesseract.exe on computer
tess.pytesseract.tesseract_cmd = r'D:\Instalation Files\tesseract\tesseract.exe'

date = '17.01.22'
# reading image from computer with the help of opencv-python or cv2
location = r'C:\Users\Vimal Panara\Pictures\Screenshots'
name_1 = location + '\{}-1.png'.format(date)
name_2 = location + '\{}-2.png'.format(date)
image_1 = cv2.imread(name_1)
image_2 = cv2.imread(name_2)

# extracting text from images of screen shot of zoom meeting
names_image_1 = tess.image_to_string(image_1).split('\n')
names_image_2 = tess.image_to_string(image_2).split('\n')

list_1 = find_names(names_image_1)
list_2 = find_names(names_image_2)

students_present = merge_two_list(list_1, list_2)
print("Total Students in the Class: {}".format(len(students_present)))
# print(students_present)
mark_attendance(students_present)
