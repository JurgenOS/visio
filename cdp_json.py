#creates dictionary without duplicates

def res_check(prev_res, ch_dict):
    res = []
    for key, value in prev_res.items():
        for sub_key in value.keys():
            res.append({key: sub_key})
    
    if ch_dict in res:                
        #print('DISCARD')
        return False
    else:
        #print('PASS') 
        return True
    

def templ_parser(templ):

    res = {}

    for node, value in templ.items():
            
        for key,item in value.items():
            if res_check(res, item):
                try:
                    res[node].update({key:item})
                except KeyError as Er:
                    res[node] = {key:item}
    return res

if __name__ == '__main__':

    template = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                   'Fa0/2': {'R6': 'Fa0/0'},
                   'Fa0/3': {'R8': 'Fa0/4'}},
            'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
            'R6': {'Fa0/0': {'R4': 'Fa0/2'},
                   'Fa0/1': {'R8': 'Fa0/3'}},
            'R7': {'Fa0/0': {'R8': 'Fa0/2'}},
            'R8': {'Fa0/2': {'R7': 'Fa0/0'},
                   'Fa0/3': {'R6': 'Fa0/1'},
                   'Fa0/4': {'R4': 'Fa0/3'}
            }}


    template2 = {'R4': {'Fa0/1': {'R5': 'Fa0/1'},
                   'Fa0/2': {'R6': 'Fa0/0'},
                   'Fa0/3': {'R8': 'Fa0/4'}},
            'R5': {'Fa0/1': {'R4': 'Fa0/1'}},
            'R7': {'Fa0/0': {'R8': 'Fa0/2'}},
            'R8': {'Fa0/2': {'R7': 'Fa0/0'},
                   'Fa0/3': {'R6': 'Fa0/1'},
                   'Fa0/4': {'R4': 'Fa0/3'}
            }}

    print(templ_parser(template))