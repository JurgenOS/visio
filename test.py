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
    
        #i = len(tire_1)
        #host = "{}".format(node)
        visio.dropShape(pagObj, stnObj.Masters.ItemU('Rectangle'), 5, i + 1, node)