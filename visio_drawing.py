import visio

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


tire_1 = {}
tire_2 = {}
tire_3 = {}



for node, value in template.items():
    
    if len(value) >= 3:
        i = len(tire_1)
        host = "{}".format(node)
        tire_1(key) = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, i + 1, host)
    if 3 > len(value) >= 2:
        i = len(tire_2)
        host = "{}".format(node)
        tire_2(key) = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 3, i + 1, host)
    if 2 > len(value) >= 1:
        i = len(tire_3)
        host = "{}".format(node)
        tire_3(key) = visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 1, i + 1, host)

nodes = {}
nodes.update(tire_1)    
nodes.update(tire_2)    
nodes.update(tire_3)    

for key, value in templ.items():
    for interf, item in value.items():
        first = nodes[key]
        second = nodes[list(item.keys())[0]]
        first_inf = 'Connections.X{}'.format(interf.split('/')[-1])
        second_inf = 'Connections.X{}'.format((list(item.values())[0]).split('/')[-1])
        note = '{}--{}'.format(first_inf, second_inf)

        visio.connectShapes2(pagObj, appVisio, first, second, first_inf, second_inf, note)
