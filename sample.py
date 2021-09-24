import gspread
from oauth2client.service_account import ServiceAccountCredentials


#used scop
scope = ['https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive']

#authentication data
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)

#authenticates
gc = gspread.authorize(credentials)

#open spreadsheet
wks = gc.open_by_key('1xg-m8v8dymLSf2WndkgBmflA4fTOUDTdJHK4ZLODbIk')

#select spreadsheet by name
wks = gc.open('Engenharia de Software – Desafio [Leonardo Sasso]') 

#select the first page of the spreadsheet
worksheet = wks.get_worksheet(0)

#reads, manipulates, and enters data into the spreadsheet



#reads, manipulates, and enters data into the spreadsheet
for line in range (4, 28): 
    absence = worksheet.cell(line, 3).value
    if int(absence) <= 15:
        first_test = worksheet.cell(line, 4).value
        second_test = worksheet.cell(line, 5).value
        third_test = worksheet.cell(line, 6).value
        final_media = (float(first_test) + float(second_test) + float(third_test))/30
        grade_for_approval = 10 - final_media
        if final_media < 5:
            worksheet.update_cell(line, 7, 'Reprovado por Nota')
            worksheet.update_cell(line, 8, '0')
        if final_media >=5 and final_media <7:
            worksheet.update_cell(line, 7, 'Exame Final')
            worksheet.update_cell(line, 8, grade_for_approval)
        if final_media > 7:
            worksheet.update_cell(line, 7, 'Aprovado')
            worksheet.update_cell(line, 8, '0')
    else:
        worksheet.update_cell(line, 7, 'Reprovado por Falta')
        worksheet.update_cell(line, 8, '0')
