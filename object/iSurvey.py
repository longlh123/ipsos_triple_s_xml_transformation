import numpy as np
import re
import xml.etree.ElementTree as ET

class iSurvey(dict):
    def __init__(self, xml_file):
        self.__dict__ = dict()
        self.openTripleS_XML(xml_file)
    
    def openTripleS_XML(self, xml_file):
        others_list = ["Khác, ghi rõ"]

        tree = ET.parse(xml_file)
        root = tree.getroot()

        survey = root.find('survey')

        self["title"] = survey.find('title').text

        record = survey.find('record')

        self["variables"] = dict()

        for variable in record.findall('variable'):
            question_name = variable.find('name').text
            
            if question_name == "Q1":
                a = ""

            if question_name not in self["variables"].keys():
                self["variables"][question_name] = dict()
                
            self["variables"][question_name]['label'] = variable.find('label').text
            self["variables"][question_name]['type'] = variable.attrib['type']
            self["variables"][question_name]['syntax'] = ""

            self["variables"][question_name]['position'] = dict()
            self["variables"][question_name]['position']['start'] = int(variable.find('position').attrib['start']) - 1
            self["variables"][question_name]['position']['finish'] = int(variable.find('position').attrib['finish'])

            match self["variables"][question_name]['type']:
                case 'quantity':
                    self["variables"][question_name]['syntax'] = "%s [py_setColumnName=%s] \"%s\" double;" % (question_name, question_name, self["variables"][question_name]['label'])
                case 'character':
                    self["variables"][question_name]['syntax'] = "%s [py_setColumnName=%s] \"%s\" text;" % (question_name, question_name, self["variables"][question_name]['label'])
                case "single" | 'multiple':
                    self["variables"][question_name]['values'] = dict()
                    self["variables"][question_name]['helperfields'] = dict()

                    values = list()
                    helperfields = list()

                    for value in variable.find('values').findall('value'):
                        if value.attrib['code'] not in self["variables"][question_name]['values'].keys():
                            self["variables"][question_name]['values'][value.attrib['code']] = dict()
                        
                        if value.text in others_list:
                            if value.attrib['code'] not in self["variables"][question_name]['helperfields'].keys():
                                self["variables"][question_name]['helperfields'][value.attrib['code']] = dict()

                        self["variables"][question_name]['values'][value.attrib['code']]['label'] = value.text
                        self["variables"][question_name]['values'][value.attrib['code']]['syntax'] = "_%s \"%s\"" % (value.attrib['code'], value.text)
                        values.append("_%s \"%s\"" % (value.attrib['code'], value.text))

                        if value.text in others_list:
                            helperfields.append(value.attrib['code'])
                    
                    helperfields_syntax = ""

                    if len(self["variables"][question_name]['helperfields']) > 0:
                        idx = 1
                        for helperfield_name, helperfield in self["variables"][question_name]['helperfields'].items():
                            helperfield["name"] = "%sr97%soe" % (question_name, "" if len(self["variables"][question_name]['helperfields']) == 1 else idx)
                            idx += 1

                        helperfields_syntax = "helperfields({})".format(";".join(["{} \"{}\" text".format(v['name'], k) for k, v in self["variables"][question_name]['helperfields'].items()])) 

                    if self["variables"][question_name]['type'] == 'single':
                        self["variables"][question_name]['syntax'] = "%s [py_setColumnName=%s] \"%s\" categorical[1..1]{%s}%s;" % (question_name, question_name, self["variables"][question_name]['label'], ",".join(values), helperfields_syntax)
                    else:
                        self["variables"][question_name]['syntax'] = "%s [py_setColumnName=%s,py_showPunchingData=True] \"%s\" categorical[1..%s]{%s}%s;" % (question_name, question_name, self["variables"][question_name]['label'], len(values), ",".join(values), helperfields_syntax)
                