from openpyxl import load_workbook

from numpy import *

import csv, sys

cube_sizes = {'Large': 24.0/2.0, 'Medium': 18.0/2.0, 'Small': 12.0/2.0}

def generate_edges(width):
	vertices = [ [1, 1, 1], [-1, 1, -1], [-1, -1, 1], [1, -1, -1] ]
	edges = []
	for vertex in vertices:
		for dimension in range(3):
			edge = {}
			second_vertex = vertex[:]
			# Swap the sign of the vertex in the current dimension
			second_vertex[dimension] = -1*second_vertex[dimension]
			edge['p1'] = [val*width for val in vertex]
			edge['p2'] = [val*width for val in second_vertex]

			edge['id'] = len(edges)

			edges.append(edge)

	# Output dictionary with keys 'p1' 'p2' 'id'
	return edges



# def load_spreadsheet(filename, sheet_name="Sheet1"):
def load_spreadsheet(filename, sheet_name="Sheet1"):
	wb = load_workbook(filename, data_only=True)
	spreadsheet = wb[sheet_name]
	return spreadsheet

def cellIsCube(c):
	return "Folding Cube Asm" in c and not 'Edge Asm' in c

def cellIsEdge(c):
	return 'Edge Asm' in c

# 
# PROCESS SPREADSHEET INTO APPROPRIATE PYTHON DATA STRUCTURE
# 

# 
def process(st, OLD_FORMAT=False):

	cubes = []
	cube = {}
	stripID = 0
	leds = []

	# Cycle through rows
	for row in xrange(2, st.get_highest_row()+1):

		if OLD_FORMAT is True:
			if bool(cube):
				cubes.append(cube)
			cube = {}
			cube['strips'] = []
			stripID = 0

			name = (st['A%d'%row].value)
			name, size = name.split(',')
			cube_id = name.split('-')[1]
			cube['cube_id'] = cube_id
			# print size

			# B through J are R
			# N through P are T in inches
			R0 = (st['B%d'%row].value, st['C%d'%row].value, st['D%d'%row].value)
			R1 = (st['E%d'%row].value, st['F%d'%row].value, st['G%d'%row].value)
			R2 = (st['H%d'%row].value, st['I%d'%row].value, st['J%d'%row].value)
			R = matrix( [ R0, R1, R2 ] )

			T = matrix( ( st['K%d'%row].value, st['L%d'%row].value, st['M%d'%row].value ) )
			cube['R'] = R
			cube['T'] = T

			# OLD CUBE CODE
			# Generate strips
			strips = generate_edges( width=cube_sizes[size] )

			# Add current Cube transform to strips
			for strip in strips:
				strip['R'] = cube['R']
				strip['T'] = cube['T']

			# Add strips to current Cube
			cube['strips'] = strips

		else:
			# If we're at the beginning of a Cube
			component = st['A'+str(row)].value

			if cellIsCube( component ):
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

				cube['cube_id'] = component.split('-')[1]


			# Collecting strip coordinates
			elif cellIsEdge( component ):
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
				# print cube
				# print cube.keys()
				cube['strips'].append(strip)



	for cubeIdx, cube in enumerate(cubes):
		strips = cube['strips']
		for strip in strips:
			p1 = array(strip['p1'])
			p2 = array(strip['p2'])
			num_leds_half_meter = 15
			for ledIdx in range (num_leds_half_meter):
				led = {}
				alpha = num_leds_half_meter-1
				p = (( alpha -float(ledIdx))/alpha)*p1 + (float(ledIdx)/alpha)*p2

				p = p*strip['R'] + strip['T']

				# print p

				p = p.tolist()[0]
				led['px'] = p[0]
				led['py'] = p[1]
				led['pz'] = p[2]
				# led['cubeID'] = cubeIdx
				led['cubeID'] = cube['cube_id']
				led['stripID'] = strip['id']
				led['ledID'] = ledIdx

				leds.append(led)
	return leds


# def process_old(st):

# 	cubes = []
# 	cube = {}
# 	stripID = 0
# 	leds = []

# 	# Cycle through rows
# 	for row in xrange(2, st.get_highest_row()+1):

# 		if bool(cube):
# 			cubes.append(cube)
# 		cube = {}
# 		cube['strips'] = []
# 		stripID = 0

# 		# B through J are R
# 		# N through P are T in inches
# 		R0 = (st['B%d'%row].value, st['C%d'%row].value, st['D%d'%row].value)
# 		R1 = (st['E%d'%row].value, st['F%d'%row].value, st['G%d'%row].value)
# 		R2 = (st['H%d'%row].value, st['I%d'%row].value, st['J%d'%row].value)
# 		R = matrix( [ R0, R1, R2 ] )

# 		T = matrix( ( st['K%d'%row].value, st['L%d'%row].value, st['M%d'%row].value ) )
# 		cube['R'] = R
# 		cube['T'] = T

# 		# OLD CUBE CODE
# 		# Generate strips
# 		strips = generate_edges()

# 		# Add current Cube transform to strips
# 		for strip in strips:
# 			strip['R'] = cube['R']
# 			strip['T'] = cube['T']

# 		# Add strips to current Cube
# 		cube['strips'] = strips


# 	for cubeIdx, cube in enumerate(cubes):
# 		strips = cube['strips']
# 		for strip in strips:
# 			p1 = array(strip['p1'])
# 			p2 = array(strip['p2'])
# 			for ledIdx in range (0,15):
# 				led = {}
# 				p = ((15.-float(ledIdx))/15.)*p1 + (float(ledIdx)/15.)*p2

# 				p = p*strip['R'] + strip['T']

# 				# print p

# 				p = p.tolist()[0]
# 				led['px'] = p[0]
# 				led['py'] = p[1]
# 				led['pz'] = p[2]
# 				led['cubeID'] = cubeIdx
# 				led['stripID'] = strip['id']
# 				led['ledID'] = ledIdx

# 				leds.append(led)
# 	return leds


# 	# set_printoptions(precision=2, suppress=True, threshold=nan)
# 	# print leds[-1]


# numpy.savetxt("Curvey Arches Structure.csv", a, delimiter=",")

def output_file(leds, outfilename):	
	with open('data/'+outfilename, 'wb') as csvfile:
		fieldnames = leds[0].keys()
		writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
		writer.writeheader()
		writer.writerows(leds)
		print "Output to ", outfilename, " COMPLETE"

def main(argv):
	user_input = raw_input("Enter spreadsheet filename ... ")

	# st = load_spreadsheet(filename=user_input, sheet_name=None)

	st = load_spreadsheet(filename='data/'+user_input)
	leds = process(st, OLD_FORMAT=True)
	# print len(leds)
	output_file( leds, user_input.split('.')[0] + ".csv" )





if __name__ == "__main__":
    main(sys.argv)
