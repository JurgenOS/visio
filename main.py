import visio
import win32com


'''
def xml_clear(file):
    res = ''
    for string in file.read().split('\n'):	
        res += string.strip()
    return res
'''
	
appVisio = win32com.client.Dispatch("Visio.Application")
appVisio.Visible =1

# I used a "Basic Diagram" template
doc = appVisio.Documents.Add("Basic Diagram.vst")
pagObj = doc.Pages.Item(1)
stnObj = appVisio.Documents("BASIC_M.vssx")
'''
first = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 1, 1, "First")
second = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 1, 3, "Second")

visio.connectShapes(pagObj, appVisio, first, second, 'first connect')

third = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, 1, "Third")
fourth = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, 3, "Fourth")

visio.connectShapes2(pagObj, appVisio, third, fourth, 'Connections.X1', 'PinX','second connect')
'''
template = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                   'Fa0/2': {'R6': 'Fa0/0'},
                   'Fa0/3': {'R8': 'Fa0/4'}},
            'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
            'R6': {'Fa0/0': {'R4': 'Fa0/2'},
                   'Fa0/1': {'R8': 'Fa0/3'}},
            'R7': {'Fa0/0': {'R8': 'Fa0/2'}},
            'R8': {'Fa0/2': {'R7': 'Fa0/0'},
                   'Fa0/3': {'R6': 'Fa0/1'},
                   'Fa0/4': {'R4': 'Fa0/3'}}}


templ = {'R4': {'Fa0/1': {'R5': 'Fa0/1'}, 'Fa0/2': {'R6': 'Fa0/0'}, 'Fa0/3': {'R8': 'Fa0/4'}}, 'R6': {'Fa0/1': {'R8': 'Fa0/3'}}, 'R7': {'Fa0/0': {'R8': 'Fa0/2'}}}

'''
node_list = []
nodes = {}


for key in template.keys():
    node_list.append(key)
	
for node in node_list:                                                       x  y
    nodes[node] = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, 1, "{}".format(node))
'''
tire_1 = {}
tire_2 = {}
tire_3 = {}



for node, value in template.items():
    
    if len(value) >= 3:
        i = len(tire_1)
        host = "{}".format(node)
        tire_1[node] = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), len(tire_1) + 2, 5, node)
    if 3 > len(value) >= 2:
        i = len(tire_2)
        host = "{}".format(node)
        tire_2[node] = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), len(tire_2) + 2, 3, node)
    if 2 > len(value) >= 1:
        i = len(tire_3)
        host = "{}".format(node)
        tire_3[node] = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), len(tire_3) + 2, 1, node)

nodes = {}
nodes.update(tire_1)    
nodes.update(tire_2)    
nodes.update(tire_3)

print(nodes)
		
for key, value in templ.items():
    for interf, item in value.items():
        first = nodes[key]
        second = nodes[list(item.keys())[0]]
        first_inf = 'Connections.X{}'.format(int(interf.split('/')[-1])+1)
        second_inf = 'Connections.X{}'.format(int((list(item.values())[0]).split('/')[-1])+1)
        note = '{}--{}'.format(interf.split('/')[-1], list(item.values())[0])
        #print()
        #print('first', first_inf)
        #print('second', second_inf)
        visio.connectShapes2(pagObj, appVisio, first, second, first_inf, second_inf, note)
        #visio.connectShapes(pagObj, appVisio, first, second, 'text')


	
'''	
	
for node, value in topology_json.items():	
	for interf, neighbor in value.items():
	    visio.connectShapes2(pagObj, appVisio, third, fourth, 'Connections.X1', 'PinX','second connect')
	
	
	
	

[('R4', {'Fa0/1': {'R5': 'Fa0/1'}, 'Fa0/2': {'R6': 'Fa0/0'}}), 
('R5',{'Fa0/1': {'R4': 'Fa0/1'}}), 
('R6', {'Fa0/0': {'R4': 'Fa0/2'}})]
	    
	
	
	
	    
#doc.SaveAs(r'')
#doc.Close()

#appVisio.Visible =0
#appVisio.Quit()
'''