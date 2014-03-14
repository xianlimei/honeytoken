import xml.etree.ElementTree as ET
from xml.etree.ElementTree import tostring, Element
import zipfile
import shutil
import os 
import sys, getopt
from contextlib import closing
import sys, getopt
from tempfile import mkstemp

import Document

class Ods(Document.Document):
    """Open Document Spreadsheet fileformat"""
    inputfile = ''
    outputfile = ''
    work_dir = ''
    bait = ''
    can_work = False;

    def __init__(self, input, output, work, bait):
       self.inputfile = input
       self.outputfile = output
       self.work_dir = work
       self.bait = bait
       self.can_work=False


    def inject_bait(self):
        found_it = False; ##use always when search for something
        graph_style_name = 'ax99';
        draw_id = 'idx9';
        cell_used_style='ce1'; ##seems not necessery
        row_used_style='ro1'; ##seems not necessery
        picture_name = 'PictureXX';
        last_table = ''

        ###-----------------modifying the STYLE.XML file-----------------------------###
        ####------------------------------------------------------------------------------------------------------####

        document_tree = ET.parse(self.work_dir+"/styles.xml")
        root = document_tree.getroot()
        #######################
        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}styles')
        for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name='Graphics']"):
            found_it = True;

        if(not found_it): ## if didn't find the element, what we are looking for, than we must make it
           element.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name':'Graphics','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name':graph_style_name}))
           for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name='Graphics']"):
               child =target
           
           child.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill':'none','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke':'none'}))
        found_it = False;

        ##write into file
        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir +"/styles.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/styles.xml", "a") as text_file:
            text_file.write(xmlstr)

       ####---------------------------END-OF-STYLE.XML----------------------------------------------------------------------------------------------###  
       ################################################################################# 

       ###-----------------modifying the CONTENT.XML file--------------------------------------------------------------------------------------------###
       ####---------------------------------------------------------------------------------------------------------------------------------------####



        document_tree = ET.parse(self.work_dir+"/content.xml")
        root = document_tree.getroot()
        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}automatic-styles')

        for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name='Graphics']"):
            found_it = True;

        if(not found_it): ## if didn't find the element, what we are looking for, than we must make it
           element.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name':'Graphics','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name':graph_style_name}))
           for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name='Graphics']"):
               child =target
           
           child.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill':'none','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke':'none'}))
           
        found_it = False;

        ##-----------------------inject into body section------------------------############

        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body')       ##office:body
        spreadsheet = element.find("{urn:oasis:names:tc:opendocument:xmlns:office:1.0}spreadsheet") ##office:spreadsheet
        table = spreadsheet.find("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table")  ##office:table

        ##insert the row into one before last, 'couse last is row means all empty row to ending
        table_row = Element('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row',{'{urn:oasis:names:tc:opendocument:xmlns:table:1.0}style-name':row_used_style});
        cell1 =ET.SubElement(table_row,'{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-cell',{'{urn:oasis:names:tc:opendocument:xmlns:table:1.0}sytle-name':cell_used_style});
        draw_frame = ET.SubElement(cell1,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}frame',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}id':draw_id,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name':picture_name,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style-name':graph_style_name,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}z-index':'1','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-height':'scale','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-width':'scale','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height':'5.53125in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width':'8.33333in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}x':'0in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y':'0in'})
        draw_image = ET.SubElement(draw_frame,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}image',{'{http://www.w3.org/1999/xlink}actuate':'onLoad','{http://www.w3.org/1999/xlink}href':self.bait,'{http://www.w3.org/1999/xlink}show':'embed','{http://www.w3.org/1999/xlink}type':'simple'})
        title = ET.SubElement(draw_frame,'{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}title')
        desc = ET.SubElement(draw_frame, '{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}desc')
        cell2 = ET.SubElement(table_row,'{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-cell',{'{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated':'16383'})



        table.insert(len(table)-1,table_row)

        
        for target in table.findall("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}table-row"):##in the last row is the nr of empty rows
            last_table = target;

        if(not (last_table =='')):
            print last_table.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated')
            if(not (last_table.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated')==None)):
                key_nr = int(last_table.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated'))
                key_nr = int(key_nr)-2
            else:
                key_nr = 1000
            last_table.attrib['{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-rows-repeated'] =str(key_nr); ##decrease the empty row nr with 2, the size of the picture




        
        ##write into file
        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir +"/content.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/content.xml", "a") as text_file:
            text_file.write(xmlstr)
        
        self.zipdir(self.work_dir, self.outputfile)  
         
        return True

    #namepsaces for odt from LibreOffice + some for MO
    ET.register_namespace('office', "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
    ET.register_namespace('style', "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
    ET.register_namespace('text', "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
    ET.register_namespace('table', "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
    ET.register_namespace('draw', "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0")
    ET.register_namespace('fo', "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0")
    ET.register_namespace('xlink', "http://www.w3.org/1999/xlink")
    ET.register_namespace('dc', "http://purl.org/dc/elements/1.1/")
    ET.register_namespace('meta', "urn:oasis:names:tc:opendocument:xmlns:meta:1.0")
    ET.register_namespace('number', "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
    ET.register_namespace('svg', "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0")
    ET.register_namespace('chart', "urn:oasis:names:tc:opendocument:xmlns:chart:1.0")
    ET.register_namespace('dr3d', "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0")
    ET.register_namespace('math', "http://www.w3.org/1998/Math/MathML")
    ET.register_namespace('form', "urn:oasis:names:tc:opendocument:xmlns:form:1.0")
    ET.register_namespace('script', "urn:oasis:names:tc:opendocument:xmlns:script:1.0")
    ET.register_namespace('ooo', "http://openoffice.org/2004/office")
    ET.register_namespace('ooow', "http://openoffice.org/2004/writer")
    ET.register_namespace('oooc', "http://openoffice.org/2004/calc")
    ET.register_namespace('dom', "http://www.w3.org/2001/xml-events")
    ET.register_namespace('xforms', "http://www.w3.org/2002/xforms")
    ET.register_namespace('xsd', "http://www.w3.org/2001/XMLSchema")
    ET.register_namespace('xsi', "http://www.w3.org/2001/XMLSchema-instance")
    ET.register_namespace('rpt', "http://openoffice.org/2005/report")
    ET.register_namespace('of', "urn:oasis:names:tc:opendocument:xmlns:of:1.2")
    ET.register_namespace('xhtml', "http://www.w3.org/1999/xhtml")
    ET.register_namespace('grddl', "http://www.w3.org/2003/g/data-view#")
    ET.register_namespace('tableooo', "http://openoffice.org/2009/table")
    ET.register_namespace('field', "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0")
    ET.register_namespace('formx', "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0")
    ET.register_namespace('css3t', "http://www.w3.org/TR/css3-text/")
    ET.register_namespace('msoxl', "http://schemas.microsoft.com/office/excel/formula")

