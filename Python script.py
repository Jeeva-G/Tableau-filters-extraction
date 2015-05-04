'''
Created on 2/05/2015

@author: jeevananthamganesan
'''
import sys
import pandas as pd
from collections import defaultdict
from xml.etree import ElementTree


def worksheetname(worksheet,x):
    d = defaultdict(dict)
    y = 0
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
                d[x]['Values used'] = [eachelem.attrib.get('member') for eachelem in k if eachelem.attrib.get('member') != None]
            elif k.attrib.get('function') == 'except':
                d[x]['Worksheet name'] = worksheet.attrib.get('name')
                d[x]['Filter on field'] = j.attrib.get('column')
                if k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration') != None:
                    d[x]['Function used'] = k.attrib.get('{http://www.tableausoftware.com/xml/user}ui-enumeration')
                else:
                    d[x]['Function used'] = k.attrib.get('function')
                if j.attrib.get('kind') != None:
                    d[x]['Kind'] = j.attrib.get('kind')
                d[x]['Values used'] = [eachelem.attrib.get('member') for eachelem in k.iter('groupfilter') if eachelem.attrib.get('member') != None]
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
    y = y + 1
    return d

def treegeneration(twbFileName):
    with open(twbFileName, 'rt') as f:
        tree = ElementTree.parse(f)
    mydf = {}
    for i in tree.findall('./worksheets'):
        x = 0
        for worksheet in i:
            out = worksheetname(worksheet,x)
            x = x + 1
            mydf.update(out)
    mydfv1 = pd.DataFrame(mydf)
    mydfv2 = mydfv1.T
    return mydfv2


if __name__ == '__main__':
    twbFileName = sys.argv[1]
    outputcsv = sys.argv[2]
    mydfv2 = treegeneration(twbFileName)
    mydfv2.to_csv(outputcsv)
