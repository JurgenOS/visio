# -*- encoding: utf-8 -*-

import bs4 as bs

soup = bs.BeautifulSoup(open('test.xml'), 'xml')


amount_of_intf = 48






#f = open('stancill_1.xml', 'w')
for intf in range(1, amount_of_intf + 1):
    #print(templ.format(intf, amount_of_intf, intf))
    Row = soup.new_tag("Row", IX = '"{}"'.format(intf))

	
    Cell_W = soup.new_tag("Cell", F = "Width/{}*{}".format(amount_of_intf, intf), N = "X", V = "0")
    Cell_H = soup.new_tag("Cell", F = "Height*0", N = "Y", V = "0")
    Cell_X = soup.new_tag("Cell", N = "DirX", V = "0")	
    Cell_Y = soup.new_tag("Cell", N = "DirY", V = "0")
    Cell_T = soup.new_tag("Cell", N = "Type", V = "0")
    Cell_Au = soup.new_tag("Cell", N = "AutoGen", V = "0")
    Cell_N_F = soup.new_tag("Cell", F="No Formula", N="Prompt", V="")
	
	
    Row.append(Cell_W)
    Row.append(Cell_H)
    Row.append(Cell_X)
    Row.append(Cell_Y)
    Row.append(Cell_T)
    Row.append(Cell_Au)
    Row.append(Cell_N_F)
	
    soup.Shapes.Shape('Section', N = 'Connection')[0].append(Row)
	
'''
    templ ='<Row IX="{}">\
<Cell F="Width/{}*{}" N="X" V="0"/>\
<Cell F="Height*0" N="Y" V="0"/>\
<Cell N="DirX" V="0"/>\
<Cell N="DirY" V="0"/>\
<Cell N="Type" V="0"/>\
<Cell N="AutoGen" V="0"/>\
<Cell F="No Formula" N="Prompt" V="0"/>\
</Row>'
    soup.Shapes.Shape('Section', N = 'Connection')[0].append(templ.format(intf, amount_of_intf, intf))
    #f.write(templ.format(intf, amount_of_intf, intf))
    print(soup.prettify())'''
	
	
	
f = open('stancill_1.xml', 'w')
f.write(str(soup))