a = 'reporte_cancelados'
b = 'reporte_precongreso_aprobados'  # talleres comprados
c = 'ReporteCongresoCanselado'
d = 'ReporteCongresoAprobados'  # estudiante y profecional aprovado


l = ['// Problemas especÃ­ficos de  aprendizaje // TÃ©cnicas de neuroimagen en  NeuropsicologÃ­a  // ',
     '// NeuropsicologÃ­a en el colegio:  mÃ¡s allÃ¡ de las pruebas // ValoraciÃ³n de Funciones Ejecutivas  y del Desarrollo de 3 a 6 aÃ±os . // NeurobiologÃ­a de la dislexia.  Del diagnÃ³stico al tratamiento  ',
     '// Problemas especÃ­ficos de  aprendizaje // EvaluaciÃ³n neuropsicolÃ³gica  del niÃ±o pre-escolar  // NeurobiologÃ­a de la dislexia.  Del diagnÃ³stico al tratamiento  ',
     '// EvaluaciÃ³n de las demencias  // TÃ©cnicas de neuroimagen en  NeuropsicologÃ­a  // NeuropsicologÃ­a forense: facilitar los derechos y la justicia para las personas con discapacidades cerebrales. '

]
'''
talleres = []
for i in range(len(l)):
	g = l[i]
	h = g.split('// ')
	#print(h, '\n @@@@')
	topop =[]
	for r in range(len(h)):
		if h[r] == '':
			topop.append(r)

	topop.sort(reverse=True)
	for q in topop:
		h.pop(q)
	print(h,'\n ######')
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

print(talleres)
'''
def numerarTalleres(tallerExcel):
	talleres = []
	g = tallerExcel
	h = g.split('// ')
	#print(h, '\n @@@@')
	topop =[]
	for r in range(len(h)):
		if h[r] == '':
			topop.append(r)

	topop.sort(reverse=True)
	for q in topop:
		h.pop(q)
	print(h,'\n ######')
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

	print(talleres)

numerarTalleres(l[3])