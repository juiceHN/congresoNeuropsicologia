from xlrd import open_workbook
t1, t2, t3, t4, t5, t6 = [], [], [], [], [], []
t7, t8, t9, t10, t11, t12 = [], [], [], [], [], []


def readDoc(filename, page):
    wb = open_workbook(filename)
    page = wb.sheet_by_index(page)
    return wb, page


def cleanLine(line):
    char = "'"
    for c in char:
        line = line.replace(c, "")
    return line


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


def fetchAllCongreso(ce, cp):
    names = getAllColumn(ce, 0, "text:'First Name'")
    lastNames = getAllColumn(ce, 0, "text:'Last Name'")
    email = getAllColumn(ce, 0, "text:'Email'")
    qt = getAllColumn(ce, 0, "text:'Ticket Type'")
    a = len(names)
    b = []
    for i in range(a):
        b.append('Eventbrite')
    return names, lastNames, email, qt, b


def numerarTalleres(tallerExcel):
    talleres = []
    g = tallerExcel
    h = g.split('// ')
    topop = []
    for r in range(len(h)):
        if h[r] == '':
            topop.append(r)

    topop.sort(reverse=True)
    for q in topop:
        h.pop(q)
    # print(h,'\n ######')
    for j in h:
        t = j.split(' ')
        print(t)
        if 'Problemas' in t:
            talleres.append(1)
        elif 'ejercicio' in t:
            talleres.append(2)
        elif 'demencias' in t:
            talleres.append(3)
        elif 'colegio:' in t:
            talleres.append(4)
        elif 'pre-escolar' in t:
            talleres.append(5)
        elif 'Desarrollo' in t:
            talleres.append(6)
        elif 'validez' in t:
            talleres.append(7)
        elif 'neuroimagen' in t:
            talleres.append(8)
        elif 'Reconocimiento' in t:
            talleres.append(9)
        elif 'dislexia.' in t:
            talleres.append(10)
        elif 'breve' in t:
            talleres.append(11)
        elif 'forense:' in t:
            talleres.append(12)
        else:
            talleres.append(0)

    return talleres


def fetchTalleres(fT):
    numTalleres = []
    names = getAllColumn(fT, 0, "text:'Nombre'")
    email = getAllColumn(fT, 0, "text:'Correo'")
    phone = getAllColumn(fT, 0, "text:'Telefono'")
    talleres = getAllColumn(fT, 0, "text:'Talleres elegidos'")
    for i in range(len(talleres)):
        new = numerarTalleres(talleres[i])
        numTalleres.append(new)
    print(numTalleres)
    return names, email, phone, numTalleres


#fetchTalleres('pruebas.xlsx')