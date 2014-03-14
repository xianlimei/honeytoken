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

class Odt(Document.Document):

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
        namestring0 = "a0"
        namestring2 = "T2"
        found =0;

        document_tree = ET.parse(self.work_dir+"/content.xml")
        root = document_tree.getroot()
        #######################
        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}automatic-styles')
        element.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name': namestring2,'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name':'DefaultParagraphFont','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'text'}))
        for styler in element.findall( '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style' ):
            if(styler.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name']==namestring2):
                child_elem =styler

        child_elem.append(Element('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}text-properties',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}language-asian':'hu', '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}country-asian':'HU'}))
        ##########################
        element.append(Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name': namestring0,'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name':'Graphics','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic'}))
        for styler in element.findall( '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style' ):
            if(styler.attrib['{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name']==namestring0):
                child_elem =styler

        child_elem.append(Element('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}graphic-properties',{'{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}border':'0.01042in none', '{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}background-color':'transparent','{urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0}clip':'rect(0in 0in 0in 0in)'}))
    
        ###########################
        body = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body')
        text= body.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}text')
        p= text.find('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p')   #name P1 by mo Standard by lo
        for spanner in p.findall( '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style' ):
            if(spanner.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name']==namestring2):
                span =spanner
                found=1

        if found == 0:
            p.append(Element('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span',{'{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name':namestring2}))
            for spanner in p.findall('{urn:oasis:names:tc:opendocument:xmlns:text:1.0}span'):
                if(spanner.attrib['{urn:oasis:names:tc:opendocument:xmlns:text:1.0}style-name']==namestring2):
                    span =spanner
        else:
            print "found it"
            found=0;

        span.append(Element('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}frame',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style-name':namestring0,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name': 'Picture 1','{urn:oasis:names:tc:opendocument:xmlns:text:1.0}anchor-type':'as-char','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}x':'0in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y':'0in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width':'3.58333in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height':'2.36458in','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-width':'scale','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-height':'scale'}))
        frame = span.find('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}frame')
        frame.append(Element('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}image',{'{http://www.w3.org/1999/xlink}href':self.bait,'{http://www.w3.org/1999/xlink}type':'simple','{http://www.w3.org/1999/xlink}show':'embed','{http://www.w3.org/1999/xlink}actuate':'onLoad'}))
        frame.append(Element('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}title'))
        frame.append(Element('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}desc'))
        desc = frame.find('{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}desc')
        desc.text = self.bait


        print("The xml manipulation is done!")
        #end of the xml manipulation
        #############################################
        xmlstr = ET.tostring(root)

        #write xml header
        with open(self.work_dir +"/content.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/content.xml", "a") as text_file:
            text_file.write(xmlstr)

        print("The decoy injection is done!")
        ############################################


        raw_input("Press any button to exit. ")
        self.zipdir(self.work_dir, self.outputfile)    


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

    #namespaces for odt from Microsoft Office
    ET.register_namespace('presentation', "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
    ET.register_namespace('smil', "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0")
