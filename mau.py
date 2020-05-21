from openpyxl import load_workbook
from datetime import datetime

workbook = load_workbook(filename="mau.xlsx")
sheet = workbook.active


matricula = 0
nombreAlumno = 1
nivel = 2
programa = 3
estatus = 4
ultimaAsistencia = 5

info = ['A', 'B', 'C', 'D', 'E', 'F']


comentario = "ssasdas"
lugar = "PEDRO MORENO"

personIndex = 2
for personIndex in range(2, 22):
    informacion = []
    for letter in info:
        if(letter == 'F'):
            informacion.append([])
            cell = letter+str(personIndex)
            dateFull = sheet[cell].value
            year = dateFull.strftime("%Y")
            month = dateFull.strftime("%m")
            day = dateFull.strftime("%d")
            informacion[5].append(year)
            informacion[5].append(day)
            informacion[5].append(month)
        else:
            cell = letter+str(personIndex)
            informacion.append(sheet[cell].value)
    print(informacion)
    comentario = input("Ingrese el comentario de la persona: ")
    print("Haciendo envio")
    
    final_url = "entry.142362158_sentinel=&entry.142362158="+str(lugar)+"&entry.2070048545=" + str(informacion[matricula])+"&entry.301730181_year="+str(informacion[5][0])+"&entry.301730181_month="+str(informacion[5][2])+"&entry.301730181_day="+str(informacion[5][1])+"&entry.2108397311_hour=08&entry.2108397311_minute=30&entry.246699160="+str(informacion[estatus])+"&entry.1073304337="+str(informacion[nivel])+"&entry.1118742320="+"OK"+"&entry.969578486="+comentario+"&fvv=1&draftResponse=%5Bnull%2Cnull%2C%22-768346404936795950%22%5D%0D%0A&pageHistory=0&fbzx=-768346404936795950"
    original_url = "https://docs.google.com/forms/d/e/1FAIpQLSdPujwS_4h0BBPyrnMsL0uOzOWLT8TXnqK7fAQn5t0SUkQWmQ/formResponse?"
    print(original_url + final_url)

