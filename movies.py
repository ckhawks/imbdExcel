import openpyxl
import requests

wb = openpyxl.load_workbook('jp.xlsx', read_only=False)

#http://www.omdbapi.com/?t=Baby+Driver
sheet = wb.get_sheet_by_name("Film")
target = wb.copy_worksheet(sheet)

for i in range(2, 268):
    name = target['A' + str(i)].value
    print(name)
    params = [('t', name), ('apikey', 'aaaaaaaa')]
    request = requests.get('http://www.omdbapi.com/', params=params)
    print(request.content)
    jsons = request.json()
    if "Movie not found" in request.text:
        print("rip")
    else:
        print("genre: " + jsons["Genre"])
        target['C' + str(i)] = jsons["Genre"]
        print("director: " + jsons["Director"])
        target['B' + str(i)] = jsons["Director"]
wb.template = False
wb.save('jpmod.xlsx')#, as_template=False)
