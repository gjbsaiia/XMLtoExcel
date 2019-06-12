#!python3
import os
import sys
import openpyxl
import xml.etree.ElementTree as ET

# data file name
wkbk = "data.xlsx"
# music sheet name
mu = "Music"
# food sheet name
fu = "Food"
# activities sheet name
act = "Activities"

# folder that houses all xmls that need to be processed
xml_fold = "raw_data/"
# all music XML
mu_xml = []
# attributes belonging to music
mu_att = ["Song"]
# all food XML
fu_xml = []
# attributes belonging to food
fu_att = ["HamB", "VegB", "HDog", "VDog", "Ketchup", "Mustard", "Mayo", "Relish", "Lettuce", "Tomato", "Onion", "Pickles"]
# all activity XML
act_xml = []
# attributes belonging to activity
act_att = ["Name", "Member2", "Tug_Y_N", "Home_Y_N", "Pin_Y_N", "Rock", "HipHop", "Alternative", "RnB", "EDM", "Pop"]

# Maps XML tag to location on Excel Sheet
# Convention: XML = ["Column", LastIndex, IndexShouldBeUpdated]
indexMap = {
    "Name": ["A",2, True],
    "Member2": ["B",2, True],
    "Tug_Y_N": ["C",2, True],
    "Home_Y_N": ["D",2, True],
    "Pin_Y_N": ["E", 2, True],
    "Rock": ["L", 2, False],
    "HipHop": ["M", 2, False],
    "Alternative": ["N", 2, False],
    "RnB": ["O", 2, False],
    "EDM": ["P", 2, False],
    "Pop": ["Q", 2, False],
    "HamB": ["A", 2, False],
    "VegB": ["B", 2, False],
    "HDog": ["C", 2, False],
    "VDog": ["D", 2, False],
    "Ketchup": ["E", 2, False],
    "Mustard": ["F", 2, False],
    "Mayo": ["G", 2, False],
    "Relish": ["H", 2, False],
    "Lettuce": ["I", 2, False],
    "Tomato": ["J", 2, False],
    "Onion": ["K", 2, False],
    "Pickles": ["L", 2, False],
    "Song": ["A", 2, True]
}

def main():
    newData()
    if mu_xml:
        sheet = getSheet(mu)
        updateIndexes(sheet, mu_att)
        values = stripMuXml(mu_xml, mu_att)
        writeExcel(sheet, values, mu_att)
    if act_xml:
        sheet = getSheet(act)
        updateIndexes(sheet, act_att)
        values = stripMuXml(act_xml, act_att)
        writeExcel(sheet, values, act_att)
    if fu_xml:
        sheet = getSheet(fu)
        values = stripMuXml(fu_xml, fu_att)
        writeFoodExcel(sheet, values, fu_att)
    
# update indexes in IndexMap
def updateIndexes(sheet, att):
    for each in att:
        if indexMap[each][2]:
            indexMap[each].update([indexMap[each][0], getMaxRow(sheet, indexMap[each][0])+1, indexMap[each][2]])

# returns sheet corresponding to specific data set
def getSheet(name):
    global wkbk
    wb = openpyxl.load_workbook(wkbk)
    sheet = wb.get_sheet_by_name(name)
    return sheet

# checks for and organizes new data sets
def newData():
    global mu_xml, fu_xml, act_xml, xml_fold
    resp = os.listdir(xml_fold)
    for each in resp:
        if "Song" in each:
            mu_xml.append(xml_fold+each)
        if "food" in each:
            fu_xml.append(xml_fold+each)
        if "activities" in each:
            act_xml.append(xml_fold+each)

# helper to return Max row cleanly
def getMaxRow(sheet, col):
    return len(sheet[col])

# wrapper method to handle XML extraction and writing data into excel
def writeExcel(sheet, values, att):
    i = 0
    for each in values:
        if(i >= len(att)):
            break
        sheet[indexMap[att[i]][0]+str(indexMap[att[i]][1])] = each
        if(indexMap[att[i]][2]):
            indexMap[att[i]].update(indexMap[att[i]][0], indexMap[att[i]][1]+1, True)

def stripActXml(xml, att):
    values = []
    tree = ET.parse(path)
    root = tree.getroot()
    xmlstr = ET.tostring(root, encoding='utf8', method='xml')
    xmlLis = xmlstr.split("\n")
    j = 0
    for each in xmlLis:
        if(">" in each and "<" in each):
            splitted = each.split(">")
            if splitted[1] and att[j] in splitted[1]:
                values.append(splitted[1].split("<")[0])
                if( j < len(att)):
                    j+=1
    return values

def stripMuXml(xml, att):
    values = []
    tree = ET.parse(path)
    root = tree.getroot()
    xmlstr = ET.tostring(root, encoding='utf8', method='xml')
    xmlLis = xmlstr.split("\n")
    j = 0
    for each in xmlLis:
        if(">" in each and "<" in each):
            splitted = each.split(">")
            if splitted[1] and "field2" in splitted[1]:
                if "," in splitted[1].split("<")[0]:
                    splittted = splitted[1].split("<")[0].split(",")
                    for each in splittted:
                        values.append(each)
                else:
                    values.append(splitted[1].split("<")[0])
                break
    return values

def stripFuXml(xml, att):
    tree = ET.parse(path)
    root = tree.getroot()
    xmlstr = ET.tostring(root, encoding='utf8', method='xml')
    xmlLis = xmlstr.split("\n")
    j = 0
    for each in xmlLis:
        next = False
        v = 1
        if(">" in each and "<" in each):
            splitted = each.split(">")
            if splitted[1]:
                if att[j] in splitted[1]:
                    values.append(splitted[1].split("<")[0])
                    if( j < len(att)):
                        j+=1
                if "Burger" in splitted[1] or "Dogs":
                    if(int(splitted[1].split("<")[0]) == 1):
                        v = -1
                    next = True
                    j+=1
                if(next):
                    values.append(int(splitted[1].split("<")[0]) * -1)
                    j+=1
    return values




if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print('Interrupted \_[*.*]_/\n')
        try:
            sys.exit(0)
        except SystemExit:
            os.exit(0)