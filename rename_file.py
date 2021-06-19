import xlrd
import os
import re
import sys


ls = ''
while ls.endswith('.xlsx') == False:
	print ('Nhập đường dẫn danh sách sinh viên lớp học phần (ex: C:\ danhsach\ list.xlsx): ')
	ls = input()
print ('Nhập đường dẫn folder chứa các file cần đổi tên (ex: C:\ danhsach\ bailam\ ): ')
url = input()

# ls = '/home/caoanh/CODE/rename/rename_file/list.xlsx'
# url = '/home/caoanh/CODE/rename/tieuluan/'
wb = xlrd.open_workbook(ls)
sheet = wb.sheet_by_index(0)
print (sheet.cell_value(1, 1))
print (sheet.nrows)
print (sheet.ncols)
print ('============================')


patterns = {
    '[àáảãạăắằẵặẳâầấậẫẩ]': 'a',
    '[đ]': 'd',
    '[èéẻẽẹêềếểễệ]': 'e',
    '[ìíỉĩị]': 'i',
    '[òóỏõọôồốổỗộơờớởỡợ]': 'o',
    '[ùúủũụưừứửữự]': 'u',
    '[ỳýỷỹỵ]': 'y'
}

#"Họ Và Tên" ==>> "Ho_Va_Ten"
def convert(text):
    output = text
    for regex, replace in patterns.items():
        output = re.sub(regex, replace, output)
        # deal with upper case
        output = re.sub(regex.upper(), replace.upper(), output)
    return output



def format(text):
	print (text)
	print ('______________________________________')
	output = ""
	for i in range(0,len(text)):
		if text[i] == ' ':
			output += '_'
		else:
			output += text[i]
	return output



#split student id from 'filename.pdf' 
def ID_student(str):
	i=0
	ID = ""
	while i < len(str):
		if str[i] == '0' and str[i+1] == '3' and str[i+2] == '0' and str[i+11].isdigit() == True:
			print (str[i])
			print (i)
			for i in range(i,i+12):
				ID += str[i]
			print(ID)
		i+=1
	return ID


#get name student from 'list.xlsx' by ID
def Get_name(ID):
	col = 1
	name = ""
	for row in range(0,sheet.nrows):
		print (row)
		if ID == sheet.cell_value(row, col):
			name = sheet.cell_value(row, col + 1) + "_" + sheet.cell_value(row, col +2)
			print (name)
			break
	return name


def main():

	for count, filename in enumerate(os.listdir(url)):
		if filename.endswith('.pdf') == True:
			file_type = '.pdf'
		if filename.endswith('.docx') == True:
			file_type = '.docx'
		ID = ID_student(filename)
		name = format(convert(Get_name(ID)))
		print (name)
		dst = ID + "-" + name + file_type
		src = url + filename
		dst = url + dst

		# rename() function will
		# rename all the files

		os.rename(src, dst)  
# Driver Code
if __name__ == '__main__':
      
    # Calling main() function
    main()
