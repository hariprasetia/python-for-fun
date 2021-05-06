import MySQLdb
import xlsxwriter
import datetime


def create():
	workbook = xlsxwriter.Workbook('users.xlsx')

	worksheet = workbook.add_worksheet()
	chart = workbook.add_chart({'type': 'column'})
	format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color':'#e32f2f', 'right': True, 'right_color': '#605c5c'})
	format_date = workbook.add_format({'num_format': 'yyyy-mm-dd'})

	client = MySQLdb.connect(host="localhost", port=3306, user="user", passwd="password", db="your_database")
	cursor = client.cursor()
	query = "select userid, nama, tanggal_lahir, tinggi_badan, berat_badan, golongan_darah, alamat, telepon from users"
	cursor.execute(query)
	client.commit()
	result = cursor.fetchall()

	userid = [row[0] for row in result]
	nama = [row[1] for row in result]
	tanggal_lahir = [row[2] for row in result]
	tinggi_badan = [row[3] for row in result]
	berat_badan = [row[4] for row in result]
	golongan_darah = [row[5] for row in result]
	alamat = [row[6] for row in result]
	telepon = [row[7] for row in result]
	col = len(tanggal_lahir)+1
	headings = ['USERID', 'NAMA', 'TANGGAL LAHIR', 'TINGGI BADAN', 'BERAT BADAN', 'GOLONGAN DARAH', 'ALAMAT', 'TELEPON']
	data = [userid, nama, tanggal_lahir, tinggi_badan, berat_badan, golongan_darah, alamat, telepon]
	worksheet.set_column(2, 4, 14)
	worksheet.write_row('A1', headings, format)
	worksheet.write_column('A2', data[0])
	worksheet.write_column('B2', data[1])
	worksheet.write_column('C2', data[2], format_date)
	worksheet.write_column('D2', data[3])
	worksheet.write_column('E2', data[4])
	worksheet.write_column('F2', data[5])
	worksheet.write_column('G2', data[6])
	worksheet.write_column('H2', data[7])
	workbook.close()

if __name__ == '__main__':
	create()
