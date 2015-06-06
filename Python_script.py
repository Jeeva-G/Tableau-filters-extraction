'''
Created on 14/05/2015

@author: jeevananthamganesan
'''
import sys
import re
import pandas as pd
from collections import defaultdict
from xml.etree import ElementTree


def worksheetname(worksheet,x):
    d = defaultdict(dict)
    for j in worksheet.iter('filter'):
        for k in j.findall('groupfilter'):
            if k.attrib.get('function') == 'filter':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('function')
                d[x]['Expression'] = k.attrib.get('expression')
            elif k.attrib.get('function') == 'member':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                d[x]['Values used'] = k.attrib.get('member')
            elif k.attrib.get('function') == 'union':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                d[x]['Values used'] = []
                for eachelem in k.iter('groupfilter'):
                    if eachelem.attrib.get('function') == 'range':
                        d[x]['Values used'] = ['Range From = ' + eachelem.attrib.get('from')+ ' to ' + eachelem.attrib.get('to')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') != 'level-members':
                        if eachelem.attrib.get('member') != None:
                            d[x]['Values used'] = [eachelem.attrib.get('member')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') == 'level-members':
                        pass
                    else:
                        d[x]['Values used'] = d[x]['Values used'] + ['Something new in union'] 
            elif k.attrib.get('function') == 'except':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Values used'] = []
                for eachelem in k.iter('groupfilter'):
                    if eachelem.attrib.get('function') == 'range':
                        d[x]['Values used'] = ['Range From = ' + eachelem.attrib.get('from')+ ' to ' + eachelem.attrib.get('to')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') != 'level-members':
                        if eachelem.attrib.get('member') != None:
                            d[x]['Values used'] = [eachelem.attrib.get('member')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') == 'level-members':
                        pass
                    else:
                        d[x]['Values used'] = d[x]['Values used'] + ['Something new in except'] 
            elif k.attrib.get('function') == 'range':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                d[x]['Values used'] = k.attrib.get('level'), '--> ' +'Range From = ' + k.attrib.get('from'), ' to ' + k.attrib.get('to')
            elif k.attrib.get('function') == 'end':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('function')
                d[x]['Expression'] = (str(k.attrib.get('end'))).upper(),'' + k.attrib.get('count'), 'records'
            else:
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                d[x]['Function used'] = k.attrib.get('function')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Details'] = k.tag,k.attrib
            x = x + 1
    return (d,x)


def sharedfilters(sharedviews,x):
    d = defaultdict(dict)
    for j in sharedviews.iter('filter'):
        for k in j.findall('groupfilter'):
            if k.attrib.get('function') == 'filter':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('function')
                d[x]['Expression'] = k.attrib.get('expression')
            elif k.attrib.get('function') == 'member':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                d[x]['Values used'] = k.attrib.get('member')
            elif k.attrib.get('function') == 'union':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                d[x]['Values used'] = []
                for eachelem in k.iter('groupfilter'):
                    if eachelem.attrib.get('function') == 'range':
                        d[x]['Values used'] = ['Range From = ' + eachelem.attrib.get('from')+ ' to ' + eachelem.attrib.get('to')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') != 'level-members':
                        if eachelem.attrib.get('member') != None:
                            d[x]['Values used'] = [eachelem.attrib.get('member')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') == 'level-members':
                        pass
                    else:
                        d[x]['Values used'] = d[x]['Values used'] + ['Something new in union'] 
            elif k.attrib.get('function') == 'except':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Values used'] = []
                for eachelem in k.iter('groupfilter'):
                    if eachelem.attrib.get('function') == 'range':
                        d[x]['Values used'] = ['Range From = ' + eachelem.attrib.get('from')+ ' to ' + eachelem.attrib.get('to')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') != 'level-members':
                        if eachelem.attrib.get('member') != None:
                            d[x]['Values used'] = [eachelem.attrib.get('member')] + d[x]['Values used']
                    elif eachelem.attrib.get('function') == 'level-members':
                        pass
                    else:
                        d[x]['Values used'] = d[x]['Values used'] + ['Something new in except'] 
            elif k.attrib.get('function') == 'range':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                d[x]['Values used'] = k.attrib.get('level'), '--> ' +'Range From = ' + k.attrib.get('from'), ' to ' + k.attrib.get('to')
            elif k.attrib.get('function') == 'end':
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Function used'] = k.attrib.get('function')
                d[x]['Expression'] = (str(k.attrib.get('end'))).upper(),'' + k.attrib.get('count'), 'records'
            else:
                d[x]['Worksheet name'] = "All worksheet"
                d[x]['Filter on field'] = j.attrib.get('column')
                d[x]['Function used'] = k.attrib.get('function')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Details'] = k.tag,k.attrib
            x = x + 1
    return (d,x)
        
        
def treegeneration(twbFileName):
    with open(twbFileName, 'rt') as f:
        tree = ElementTree.parse(f)
    mydf = {}
    for i in tree.findall('./worksheets'):
        x = 0
        for worksheet in i:
            out,x = worksheetname(worksheet,x)
            x = x + 1
            mydf.update(out)
    mydf1 = {}
    x = 0
    for one in tree.findall('./shared-views'):
        out1, y = sharedfilters(one,x)
        y = y + 1
        mydf1.update(out1)
    
    mydf.update(mydf1)
    mydfv1 = pd.DataFrame(mydf)
    mydfv2 = mydfv1.T
    columns = ['Worksheet name','Filter on field','Function used','Kind','Values used']
    mydfv2 = mydfv2[columns] 
    return mydfv2


if __name__ == '__main__':
    twbFileName = sys.argv[1]
    outputcsv = sys.argv[2]
    mydfv2 = treegeneration(twbFileName)
    mydfv2['Filter on field'].replace(to_replace= [r"\[.*?\].",r"none:",r":nk",r":"], value="",inplace = True, regex=True)
    mydfv2.to_csv(outputcsv,index=False)
