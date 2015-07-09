from openpyxl import load_workbook

from numpy import *

import csv

wb = load_workbook('Dreamforce V6.xlsx', data_only=True)
st = wb['Sheet1']

cubes = []
cube = {}
stripID = 0
leds = []

# Cycle through rows
for row in range(2, st.get_highest_row()+1):
# for row in range(2, 48):

	# If we're at the beginning of a Cube
	if ("Folding Cube Asm" in st['A'+str(row)].value) and not ('Edge Asm' in st['A'+str(row)].value):
		if bool(cube):
			cubes.append(cube)
		cube = {}
		cube['strips'] = []
		stripID = 0
		# B through J are R
		# N through P are T in inches
		R0 = (st['B%d'%row].value, st['C%d'%row].value, st['D%d'%row].value)
		R1 = (st['E%d'%row].value, st['F%d'%row].value, st['G%d'%row].value)
		R2 = (st['H%d'%row].value, st['I%d'%row].value, st['J%d'%row].value)
		R = matrix( [ R0, R1, R2 ] )

		T = matrix( ( st['K%d'%row].value, st['L%d'%row].value, st['M%d'%row].value ) )
		cube['R'] = R
		cube['T'] = T

		# cube['id'] = 

	# Colleting strip coordinates
	elif 'Edge Asm' in st['A'+str(row)].value:
		strip = {}

		R0 = (st['B%d'%row].value, st['C%d'%row].value, st['D%d'%row].value)
		R1 = (st['E%d'%row].value, st['F%d'%row].value, st['G%d'%row].value)
		R2 = (st['H%d'%row].value, st['I%d'%row].value, st['J%d'%row].value)
		R = matrix( [ R0, R1, R2 ] )

		T = matrix( ( st['K%d'%row].value, st['L%d'%row].value, st['M%d'%row].value ) )
		strip['R'] = R
		strip['T'] = T

		strip['p1'] = [ st['N%d'%row].value, st['O%d'%row].value, st['P%d'%row].value ]
		strip['p2'] = [ st['Q%d'%row].value, st['R%d'%row].value, st['S%d'%row].value ]
		strip['id'] = (st['A'+str(row)].value).split('-')[-1]
		# print 'stripID = ', (st['A'+str(row)].value).split('-')[-1]
		cube['strips'].append(strip)
		# stripID+=1


# print cubes[0]
# ls = range(1,23)
# z = zeros( (1, 3) )

# cubePoints = []

# for l in ls:
# 	led = (0, 0, -l)
# 	cubePoints.append(led)

# for l in ls:
# 	led = (0, l, 0)
# 	cubePoints.append(led)

# for l in ls:
# 	led = (l, 0, 0)
# 	cubePoints.append(led)


# # cubePoints = array( [ (0, 0, 1), (0, 1, 0), (1, 0, 0) ] )

# points = empty(3)

for cubeIdx, cube in enumerate(cubes):
	strips = cube['strips']
	for strip in strips:
		p1 = array(strip['p1'])
		p2 = array(strip['p2'])
		for ledIdx in range (0,15):
			led = {}
			p = ((15.-float(ledIdx))/15.)*p1 + (float(ledIdx)/15.)*p2

			p = p*strip['R'] + strip['T']

			# print p

			p = p.tolist()[0]
			led['px'] = p[0]
			led['py'] = p[1]
			led['pz'] = p[2]
			led['cubeID'] = cubeIdx
			led['stripID'] = strip['id']
			led['ledID'] = ledIdx

			leds.append(led)



# set_printoptions(precision=2, suppress=True, threshold=nan)
# print leds[-1]

with open('Dreamforce V6.csv','wb') as csvfile:
	fieldnames = leds[0].keys()
	writer = csv.DictWriter(csvfile, fieldnames = fieldnames)

	writer.writeheader()
	writer.writerows(leds)

# numpy.savetxt("Curvey Arches Structure.csv", a, delimiter=",")