import xlrd
import os


ls = ''
while ls.endswith('.xlsx') == False:
	print ('Nhập đường dẫn danh sách sinh viên lớp học phần (ex: C:\ danhsach\ list.xlsx): ')
	ls = input()
print ('Nhập đường dẫn folder chứa các file cần đổi tên (ex: C:\ danhsach\ bailam\ ): ')
url = input()
wb = xlrd.open_workbook(ls)
sheet = wb.sheet_by_index(0)
print (sheet.cell_value(1, 1))
print (sheet.nrows)
print (sheet.ncols)
print ('============================')

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
			name = sheet.cell_value(row, col + 1) + " " + sheet.cell_value(row, col +2)
			print (name)
			break
	return name


def main():
	for count, filename in enumerate(os.listdir(url)):

		# check file type
		if filename.endswith('.pdf') == True:
			file_type = '.pdf'
		if filename.endswith('.docx') == True:
			file_type = '.docx'

		ID = ID_student(filename)
		name = Get_name(ID)
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
