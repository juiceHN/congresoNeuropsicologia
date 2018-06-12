import xlwt
import read as r
talleres = ['Problemas específicos de aprendizaje', 'Efectos del ejercicio físico en las funciones ejecutivas', 'Evaluación de las demencias', 'Neuropsicología en el colegio: más allá de las pruebas', 'Evaluación neuropsicológica del niño pre-escolar', 'Valoración de Funciones Ejecutivas y del Desarrollo de 3 a 6 años', 'La validez ecológica de la Evaluación Funcional',
            'Técnicas de neuroimagen en Neuropsicología', 'Reconocimiento cerebral del engaño', 'Neurobiología de la dislexia. Del diagnóstico al tratamiento', 'Evaluación neuropsicológica breve del niño y del adulto', 'Neuropsicología forense: facilitar los derechos y la justicia para las personas con discapacidades cerebrales.']
talleres2 = ['taller 1', 'taller 2', 'taller 3', 'taller 4', 'taller 5', 'taller 6',
             'taller 7', 'taller 8', 'taller 9', 'taller 10', 'taller 11', 'taller 12']
t1, t2, t3, t4, t5, t6 = [], [], [], [], [], []
t7, t8, t9, t10, t11, t12 = [], [], [], [], [], []
pc = 'reporte_cancelados.xlsx'
pa = 'reporte_precongreso_aprobados.xlsx'  # talleres comprados
cc = 'ReporteCongresoCancelados.xlsx'
ca = 'ReporteCongresoAprobados.xlsx'  # estudiante y profecional aprovado


def multilen():
    a = len(t1)
    b = len(t2)
    c = len(t3)
    d = len(t4)
    e = len(t5)
    f = len(t6)
    g = len(t7)
    h = len(t8)
    i = len(t9)
    j = len(t10)
    k = len(t11)
    l = len(t12)
    tot = a + b + c + d + e + f + g + h + i + j + k + l
    return [a, b, c, d, e, f, g, h, i, j, k, l], tot


def writeSummary(master, lenghts, tot, cd):
    header = ['No.', 'taller', 'codigo', 'No. Personas']
    header2 = ['No.', 'Tipo', 'No. Personas']
    congresos = ['Estudiantes', 'Profesionales', 'no respondio', 'total']
    sheet = master.add_sheet('Resumen')
    sheet.write(0, 0, 'Resumen Pre-congreso')
    for i in range(len(header)):
        sheet.write(1, i, header[i])
    for i in range(len(talleres)):
        sheet.write(i + 2, 0, i + 1)
        sheet.write(i + 2, 1, talleres[i])
        sheet.write(i + 2, 2, talleres2[i])
        sheet.write(i + 2, 3, lenghts[i])
    sheet.write(14, 2, 'Total')
    sheet.write(14, 3, tot)
    for i in range(len(header2)):
        sheet.write(17, i, header2[i])
    sheet.write(16, 0, 'Resumen Congreso')
    for i in range(len(congresos)):
        sheet.write(i + 18, 0, i + 1)
        sheet.write(i + 18, 1, congresos[i])
    for i in range(len(cd)):
        sheet.write(i + 18, 2, cd[i])
    sheet.write(21, 2, (cd[0] + cd[1] + cd[2]))
    zz = sheet.col(1)
    zz.width = 256 * 55


def writeAll():
    header3 = ['No.', 'Nombre', 'Correo', 'Telefono', 'Pais']
    masterF = xlwt.Workbook()
    names, mail, phone, talleres3 = r.fetchTalleres(pa)
    names2, phone2, mail2, ep, conts = r.fetchAllCongreso(ca)
    for i in range(len(talleres3)):
        a = [names[i], mail[i], phone[i]]
        if 1 in talleres3[i]:
            t1.append(a)
        if 2 in talleres3[i]:
            t2.append(a)
        if 3 in talleres3[i]:
            t3.append(a)
        if 4 in talleres3[i]:
            t4.append(a)
        if 5 in talleres3[i]:
            t5.append(a)
        if 6 in talleres3[i]:
            t6.append(a)
        if 7 in talleres3[i]:
            t7.append(a)
        if 8 in talleres3[i]:
            t8.append(a)
        if 9 in talleres3[i]:
            t9.append(a)
        if 10 in talleres3[i]:
            t10.append(a)
        if 11 in talleres3[i]:
            t11.append(a)
        if 12 in talleres3[i]:
            t12.append(a)
    lenghts, tot = multilen()
    writeSummary(masterF, lenghts, tot, conts)
    for i in range(len(talleres2)):
        tempsheet = masterF.add_sheet(talleres2[i])
        tempsheet.write(0, 0, talleres[i])
        w1 = tempsheet.col(1)
        w2 = tempsheet.col(2)
        w3 = tempsheet.col(3)
        w4 = tempsheet.col(4)
        w1.width = 256 * 25
        w2.width = 256 * 25
        w3.width = 256 * 25
        w4.width = 256 * 25
        for j in range(len(header3)):
            tempsheet.write(1, j, header3[j])
        if i == 0:
            for j in range(len(t1)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t1[j][0])
                tempsheet.write(2 + j, 2, t1[j][1])
                tempsheet.write(2 + j, 3, t1[j][2])
        if i == 1:
            for j in range(len(t2)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t2[j][0])
                tempsheet.write(2 + j, 2, t2[j][1])
                tempsheet.write(2 + j, 3, t2[j][2])
        if i == 2:
            for j in range(len(t3)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t3[j][0])
                tempsheet.write(2 + j, 2, t3[j][1])
                tempsheet.write(2 + j, 3, t3[j][2])
        if i == 3:
            for j in range(len(t4)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t4[j][0])
                tempsheet.write(2 + j, 2, t4[j][1])
                tempsheet.write(2 + j, 3, t4[j][2])
        if i == 4:
            for j in range(len(t5)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t5[j][0])
                tempsheet.write(2 + j, 2, t5[j][1])
                tempsheet.write(2 + j, 3, t5[j][2])
        if i == 5:
            for j in range(len(t6)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t6[j][0])
                tempsheet.write(2 + j, 2, t6[j][1])
                tempsheet.write(2 + j, 3, t6[j][2])
        if i == 6:
            for j in range(len(t7)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t7[j][0])
                tempsheet.write(2 + j, 2, t7[j][1])
                tempsheet.write(2 + j, 3, t7[j][2])
        if i == 7:
            for j in range(len(t8)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t8[j][0])
                tempsheet.write(2 + j, 2, t8[j][1])
                tempsheet.write(2 + j, 3, t8[j][2])
        if i == 8:
            for j in range(len(t9)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t9[j][0])
                tempsheet.write(2 + j, 2, t9[j][1])
                tempsheet.write(2 + j, 3, t9[j][2])
        if i == 9:
            for j in range(len(t10)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t10[j][0])
                tempsheet.write(2 + j, 2, t10[j][1])
                tempsheet.write(2 + j, 3, t10[j][2])
        if i == 10:
            for j in range(len(t11)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t11[j][0])
                tempsheet.write(2 + j, 2, t11[j][1])
                tempsheet.write(2 + j, 3, t11[j][2])
        if i == 11:
            for j in range(len(t12)):
                tempsheet.write(2 + j, 0, j + 1)
                tempsheet.write(2 + j, 1, t12[j][0])
                tempsheet.write(2 + j, 2, t12[j][1])
                tempsheet.write(2 + j, 3, t12[j][2])
    tempsheet = masterF.add_sheet('Congreso A')
    headerc = ['No', 'Nombre', 'Telefono', 'Correo', 'Actividad']
    tempsheet.write(0, 0, 'Congreso Aprobado')
    w1 = tempsheet.col(1)
    w2 = tempsheet.col(2)
    w3 = tempsheet.col(3)
    w4 = tempsheet.col(4)
    w1.width = 256 * 25
    w2.width = 256 * 25
    w3.width = 256 * 25
    w4.width = 256 * 25
    for i in range(len(headerc)):
        tempsheet.write(1, i, headerc[i])
    for i in range(len(names2)):
        tempsheet.write(2 + i, 0, i)
        tempsheet.write(2 + i, 1, names2[i])
        tempsheet.write(2 + i, 2, phone2[i])
        tempsheet.write(2 + i, 3, mail2[i])
        tempsheet.write(2 + i, 4, ep[i])

    namesC, phoneC, mailC, epC, contsC = r.fetchAllCongreso(cc)
    tempsheet = masterF.add_sheet('Congreso C')
    w1 = tempsheet.col(1)
    w2 = tempsheet.col(2)
    w3 = tempsheet.col(3)
    w4 = tempsheet.col(4)
    w1.width = 256 * 25
    w2.width = 256 * 25
    w3.width = 256 * 25
    w4.width = 256 * 25
    tempsheet.write(0, 0, 'Congreso Cancelado')
    for i in range(len(headerc)):
        tempsheet.write(1, i, headerc[i])
    for i in range(len(namesC)):
        tempsheet.write(2 + i, 0, i)
        tempsheet.write(2 + i, 1, namesC[i])
        tempsheet.write(2 + i, 2, phoneC[i])
        tempsheet.write(2 + i, 3, mailC[i])
        tempsheet.write(2 + i, 4, epC[i])
    headerTC = ['No.', 'Nombre', 'Correo', 'Telefono',
                'Taller 1', 'Taller 2', 'Taller 3']
    namesTC, mailTC, phoneTC, talleres3TC = r.fetchTalleres2(pc)
    tempsheet = masterF.add_sheet('Talleres C')
    for j in range(len(headerTC)):
        tempsheet.write(1, j, headerTC[j])
    for i in range(len(namesTC)):
        tempsheet.write(i + 2, 0, i)
        tempsheet.write(i + 2, 1, namesTC[i])
        tempsheet.write(i + 2, 2, mailTC[i])
        tempsheet.write(i + 2, 3, phoneTC[i])
        for g in range(len(talleres3TC[i])):
            tempsheet.write(i + 2, 4 + g, talleres3TC[i][g])
    w1 = tempsheet.col(1)
    w2 = tempsheet.col(2)
    w3 = tempsheet.col(3)
    w4 = tempsheet.col(4)
    w5 = tempsheet.col(5)
    w6 = tempsheet.col(6)
    w1.width = 256 * 25
    w2.width = 256 * 25
    w3.width = 256 * 25
    w4.width = 256 * 60
    w5.width = 256 * 60
    w6.width = 256 * 60
    masterF.save('MasterTest.xls')

writeAll()

'''
def writeTalleres(fT):
    names, mail, phone, talleres = r.fetchTalleres(fT)
'''
