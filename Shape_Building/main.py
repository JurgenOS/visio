# -*- encoding: utf-8 -*-

import bs4 as bs

amount_of_intf = 4

group = bs.BeautifulSoup(open('Shape_Building/Picture_sw.xml'), 'xml')

for intf in range(amount_of_intf + 1):

    Row = group.new_tag("Row", IX = '"{}"'.format(intf))
	
    Cell_W = group.new_tag("Cell", F = "Width/{}*{}".format(amount_of_intf, intf), N = "X", V = "0")
    Cell_H = group.new_tag("Cell", F = "Height*0", N = "Y", V = "0")
    Cell_X = group.new_tag("Cell", N = "DirX", V = "0")	
    Cell_Y = group.new_tag("Cell", N = "DirY", V = "0")
    Cell_T = group.new_tag("Cell", N = "Type", V = "0")
    Cell_Au = group.new_tag("Cell", N = "AutoGen", V = "0")
    Cell_N_F = group.new_tag("Cell", F="No Formula", N="Prompt", V="")
	
    Row.append(Cell_W)
    Row.append(Cell_H)
    Row.append(Cell_X)
    Row.append(Cell_Y)
    Row.append(Cell_T)
    Row.append(Cell_Au)
    Row.append(Cell_N_F)
	
    group.Shapes.Shape('Section', N = 'Connection')[0].append(Row)

def xml_clear(file):
    res = ''
    with open(file) as f:
        file_name = f.name
        for string in f.read().split('\n'):	
            res += string.strip()
		
    with open(file_name+'_clear', 'w') as f:
        f.write(res)
	
	
output_file = 'stancill_1.xml'
f = open(output_file, 'w')
f.write(str(group))
f.close()
xml_clear(output_file)