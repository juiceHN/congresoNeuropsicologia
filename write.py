import xlwt
import read as r
talleres = ['Problemas específicos de aprendizaje', 'Efectos del ejercicio físico en las funciones ejecutivas', 'Evaluación de las demencias', 'Neuropsicología en el colegio: más allá de las pruebas', 'Evaluación neuropsicológica del niño pre-escolar', 'Valoración de Funciones Ejecutivas y del Desarrollo de 3 a 6 años', 'La validez ecológica de la Evaluación Funcional',
            'Técnicas de neuroimagen en Neuropsicología', 'Reconocimiento cerebral del engaño', 'Neurobiología de la dislexia. Del diagnóstico al tratamiento', 'Evaluación neuropsicológica breve del niño y del adulto', 'Neuropsicología forense: facilitar los derechos y la justicia para las personas con discapacidades cerebrales.']
talleres2 = ['taller 1', 'taller 2', 'taller 3', 'taller 4', 'taller 5', 'taller 6',
             'taller 7', 'taller 8', 'taller 9', 'taller 10', 'taller 11', 'taller 12']


def writeSummary(master):
    header = ['No.', 'taller', 'codigo', 'No. Personas']
    header2 = ['No.', 'Tipo', 'No. Personas']
    congresos = ['Estudiantes', 'Profecionales', 'Ambos']
    sheet = master.add_sheet('Resumen')
    sheet.write(0, 0, 'Resumen Pre-congreso')
    for i in range(len(header)):
        sheet.write(1, i, header[i])
    for i in range(len(talleres)):
        sheet.write(i + 2, 0, i + 1)
        sheet.write(i + 2, 1, talleres[i])
        sheet.write(i + 2, 2, talleres2[i])
    for i in range(len(header2)):
        sheet.write(15, i, header2[i])
    for i in range(len(congresos)):
        sheet.write(i + 16, 0, i + 1)
        sheet.write(i + 16, 1, congresos[i])


def writeAll():
    masterF = xlwt.Workbook()
    writeSummary(masterF)
    for i in range(len(talleres2)):
        tempsheet = masterF.add_sheet(talleres2[i])
        tempsheet.write(0, 0, talleres[i])
    masterF.save('MasterTest.xls')

#writeAll()

def writeTalleres(fT):
    names, mail, phone, talleres = r.fetchTalleres(fT)

