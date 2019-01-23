from statistics import mean
from xlsxwriter import Workbook
import os

def best_fit(x,y):
	multi = lambda x, y : [x * y for x,y in zip(x,y)]	
	m = ((mean(x)*mean(y) - mean(multi(x,y))) / ((mean(x)**2) - mean(multi(x,x))))
	return round(m,3)

def anter(fname):
	
	#read through 'text' file
	with open(fname) as f:
		content = [x.split() for x in f.readlines()] 
	
	#Store document title and material title
	title = content[6][3]
	material = " ".join([x for x in content[4][3:]])
	
	#Clean up material title
	if "THERMAL EXPANSION" in material:
		material = material.replace("THERMAL EXPANSION ","")
	
	savnam = "{0}.xlsx".format(title)
	
	#Storing temp and expansion values
	temp_column = [float(n[1]) for n in content[16:]]
	expansion_column = [(float(n[3]) / 0.000001) for n in content[16:]]
	
	max_temp_idx = temp_column.index(max(temp_column))
	temp_index = []
	
	desired_temps = [1000,1260,1538]
	
	for temp in desired_temps:
	#In case, machinery fails halfway through
		try:
			temp_index.append([i + 5 for i, x in enumerate(temp_column[:max_temp_idx]) if (temp - x) > 0 and (temp - x) < 1][0])
		except IndexError:
			temp_index.append(max_temp_idx)
	
	temp_1000 , temp_1260 , temp_1538 = temp_index 
	
	workbook = Workbook(savnam)
	worksheet = workbook.add_worksheet()
	chartsheet = workbook.add_chartsheet()
	deg = u"\u2103"
	
	xl_values = {
				'A1':'Temperature','B1':'Expansion',
				'A2' : '(' + deg + ')','B2' : '(ppm)',
				'E4' : 'Expansion','F4' : 'Temperature',
				'G4' : 'Row',
				'F5' : desired_temps[0],'F6' : desired_temps[1],'F7' : desired_temps[2],
				'G5' : temp_1000,'G6' : temp_1260,'G7' : temp_1538
				}
				
	for key, val in xl_values.items():
		worksheet.write(key,val)
	
	worksheet.set_column('A:A', 11)
	
	worksheet.write_column('A4', temp_column)
	worksheet.write_column('B4', expansion_column)
	
	worksheet.set_column('F:F', 11)
	worksheet.write_formula('E5','=ROUND(SLOPE(INDIRECT("B4:B"&G5), INDIRECT("A4:A"&G5)),3)')
	worksheet.write_formula('E6','=ROUND(INDIRECT("B"&G6)/10000,3)')
	worksheet.write_formula('E7','=ROUND(INDIRECT("B"&G7)/10000,3)')
	
	chart1 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
	
	#Getting slope for the footer
	expansion_slope = best_fit(temp_column[:temp_1000 - 3],expansion_column[:temp_1000 - 3])
	
	foot = "THERMAL EXPANSION = {0} E-06".format(expansion_slope)
	chart1.add_series({
		'name': '=Sheet1!$B$1',
		'categories': '=Sheet1!$A$4:$A$4300',
		'values': '=Sheet1!$B$4:$B$4300',
	})
	chart1.set_title ({'name': material,
      'name_font': {
        'name': 'Arial',
        'size': 12,
    }})
	chart1.set_y_axis({'name': 'Expansion(ppm)'})
	chart1.set_x_axis({'name': 'Temperature('+ deg +')','min': 0,'max': 1600,
	'label_position': 'low'
		,'major_gridlines': {
        'visible': True}})
	chart1.set_legend({'none': True})
	chart1.set_style(1)
	chartsheet.set_header('&L&D &R&F')
	chartsheet.set_footer('&R'+ foot)
	chartsheet.set_chart(chart1)
	chartsheet.activate()
	workbook.close()
	print(material)
	print(title)
	

for root, dirs, files in os.walk(os.getcwd()):
	for file in files:
		file = os.path.join(root,file)
		if file.endswith(".A1A"):
			anter(file)
		elif file.endswith(".A2A"):
			anter(file)
