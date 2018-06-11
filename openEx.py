from xlrd import open_workbook
import xlwt


def cleanLine(line):
    char = "'"
    for c in char:
        line = line.replace(c, "")
    return line


def limpiarTalleres(array):
    r = []
    for coment in array:
        newA = coment.split('//')
        r.append(newA)
    for i in range(len(r)):
        r[i].pop(0)
    for i in range(len(r)):
        if len(r[i]) == 1:
            r[i].append('no')
            r[i].append('no')
        elif len(r[i]) == 2:
            r[i].append('no')
    return r


def readDoc(filename, page):
    wb = open_workbook(filename)
    page = wb.sheet_by_index(page)
    return wb, page


def getAllColumn(filename, page, cName):
    a = []
    doc, page = readDoc(filename, page)
    cols = page.ncols
    rows = page.nrows
    col = 1000
    for i in range(cols):
        if str(page.cell(0, i)) == cName:
            col = i
    for j in range(rows):
        text = str(page.cell(j, col))
        text = cleanLine(text)
        parts = text.split(':')
        a.append(parts[1])
    a.pop(0)
    return a


def dataPrep(fT):
    names = getAllColumn(fT, 0, "text:'First Name'")
    lastNames = getAllColumn(fT, 0, "text:'Last Name'")
    email = getAllColumn(fT, 0, "text:'Email'")
    qt = getAllColumn(fT, 0, "text:'Ticket Type'")
    return names, lastNames, email, qt


def compareEmail(correo1):
    correo3, correoNo = [], []
    correo2 = getAllColumn('Talleres.xlsx', 0, "text:'Correo'")
    for i in range(len(correo1)):
        correo3.append('no esta')
    for i in range(len(correo1)):
        if correo1[i] in correo2:
            a = correo1.index(correo1[i])
            correo3[a] = correo1[i]
        else:
            correoNo.append(correo1[i])
    print(correo3)



p = getAllColumn('Talleres.xlsx', 0, "text:'Talleres elegidos'")
talleres = limpiarTalleres(p)
print(talleres)


def writeDataT(fN1, fN2, fN3, fTarget):
    header = ['Nombre', 'Apellido', 'Correo',
              'Cantidad Talleres', 'Taller 1', 'Taller 2', 'Taller 3']
    nt, lt, et, qtt = [], [], [], []
    n1, l1, e1, qt1 = dataPrep(fN1)
    n2, l2, e2, qt2 = dataPrep(fN2)
    n3, l3, e3, qt3 = dataPrep(fN3)
    nt = n1 + n2 + n3
    lt = l1 + l2 + l3
    et = e1 + e2 + e3
    qtt = qt1 + qt2 + qt3
    print('''


        ##############################
        correos


        ''')
    compareEmail(et)
    print('''


    ##############################
    correos


    ''')
    masterF = xlwt.Workbook()
    sheet = masterF.add_sheet('Talleres')
    for i in range(len(nt)):
        sheet.write(i + 1, 0, nt[i])
        sheet.write(i + 1, 1, lt[i])
        sheet.write(i + 1, 2, et[i])
        sheet.write(i + 1, 3, qtt[i])
    for i in range(len(header)):
        sheet.write(0, i, header[i])
    masterF.save('MasterTest.xls')


writeDataT('1T.xlsx', '2T.xlsx', '3T.xlsx', 3)

