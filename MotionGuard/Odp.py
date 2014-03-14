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


class Odp(Document.Document):
    """Open Document Presentation"""
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
        """Make the xml structure and inject the choosen baitfile"""
        found_it=False
        found_it2=False
        style_name= "ax23"
        draw_id = "idx9"
        picture_name = "Picturex9"

        #######################################################################
        ######## MOdifying the styles.xml file#################################
        #######################################################################

        document_tree = ET.parse(self.work_dir +"/styles.xml")
        root = document_tree.getroot()
        #######################
        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}styles') #element = style
        for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}default-style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family='graphic']"):
            found_it = True;

        if(found_it):
            for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name='Graphics']"):
                found_it2 = True;

        if(found_it and not found_it2):
            element.insert(len(element)-1,Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name':'Graphics', '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic'}))

        if(not found_it and not found_it2):
            element.insert(len(element),Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name':'Graphics', '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic'}))
            default_style = Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}default-style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic'});
            graphic_properties =ET.SubElement(default_style,'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill':'solid','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill-color':'#4f81bd','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}opacity':'100%','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke':'solid','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-width':'0.02778in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-color':'#385d8a','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}stroke-opacity':'100%'})
            element.insert(len(element),graphic_properties)

        found_it = False
        found_it2 = False

        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/styles.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/styles.xml", "a") as text_file:
            text_file.write(xmlstr)

        #######################################################################
        ######## STYLES.XML END---------------#################################
        #######################################################################
        #######################################################################
        ######## MOdifying the content.xml file#################################
        #######################################################################

        document_tree = ET.parse(self.work_dir + "/content.xml")
        root = document_tree.getroot()
        #######################
        element = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}automatic-styles')
        for target in element.findall("./{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style[@{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name='Graphics']"):
            style_name = target.attrib.get('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name')
            found_it = True;

        if(not found_it):
            print 'Nincs Graphics'
            style = Element('{urn:oasis:names:tc:opendocument:xmlns:style:1.0}style',{'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}parent-style-name':'Graphics', '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}family':'graphic', '{urn:oasis:names:tc:opendocument:xmlns:style:1.0}name':style_name})
            graphic_properties =ET.SubElement(style,'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}graphic-properties',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}fill':'none','{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}stroke':'none'})
            element.insert(len(element),style)
            print 'elvileg beszurtuk a Grapicsot'

        found_it = False
        found_it2 = False

        draw_frame = Element('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}frame',{'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}id':draw_id,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}name':picture_name,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}style-name':style_name,'{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-height':'scale','{urn:oasis:names:tc:opendocument:xmlns:style:1.0}rel-width':'scale','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}height':'5.53125in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}width':'8.33333in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}x':'0in','{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}y':'0in'})
        draw_image = ET.SubElement(draw_frame,'{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}image',{'{http://www.w3.org/1999/xlink}href':self.bait,'{http://www.w3.org/1999/xlink}type':'simple','{http://www.w3.org/1999/xlink}show':'embed','{http://www.w3.org/1999/xlink}actuate':'onLoad'})
        svg_title = ET.SubElement(draw_frame,'{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}title')
        svg_desc = ET.SubElement(draw_frame,'{urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0}desc')

        body = root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body')
        presentation = body.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}presentation')
        for target in presentation.findall('{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}page'):
            draw_page = target

        draw_page.append(draw_frame)



        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/content.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/content.xml", "a") as text_file:
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

    ET.register_namespace('dom', "http://www.w3.org/2001/xml-events")

    ET.register_namespace('presentation', "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0")
    ET.register_namespace('smil', "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0")
