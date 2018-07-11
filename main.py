import visio
import win32com


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


node_list = []

for key in template.keys():
    node_list.append(key)
	
for node in node_list:
    visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, 1, "{}".format(node))
	


	
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