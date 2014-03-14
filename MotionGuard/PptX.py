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


class PptX(Document.Document):
    """M$ Office Presentation"""
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
        file_nr =0
        relationship_id = "rIdf"
        inSlide_id = 1

        file_nr =  str(len([name for name in os.listdir(self.work_dir + "/ppt/slides/_rels/") if name.endswith('.rels')])) #nr of files, its necesserly, because we want to inject into the last slide


        #######################################################################
        ######## MOdifying the /ppt/slides/_rels/slide[].xml.rels file#########
        #######################################################################

        document_tree = ET.parse(self.work_dir + "/ppt/slides/_rels/slide" + file_nr + ".xml.rels")
        root = document_tree.getroot()


        relationship_id = "rId"+str(len(root)+1) #it's valid, not like the old doxc example

        root.append(Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',{'Id':relationship_id,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image','Target':self.bait,'TargetMode':'External'}))



        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/ppt/slides/_rels/slide" + file_nr + ".xml.rels", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/ppt/slides/_rels/slide" + file_nr + ".xml.rels", "a") as text_file:
            text_file.write(xmlstr)

        self.replace(self.work_dir + "/ppt/slides/_rels/slide" + file_nr + ".xml.rels",'Relationships:','')
        self.replace(self.work_dir + "/ppt/slides/_rels/slide" + file_nr + ".xml.rels",'xmlns:Relationships','xmlns')

        #######################################################################
        ######## END of   /ppt/slides/_rels/slide[].xml.rels file##############
        #######################################################################

        ##################ppt/slides/slideNR.xml #########################
        document_tree = ET.parse(self.work_dir + "/ppt/slides/slide" + file_nr + ".xml")
        root = document_tree.getroot()

        for target in root.findall("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"):
            p_cSld = target

        for target in p_cSld.findall("{http://schemas.openxmlformats.org/presentationml/2006/main}spTree"):
            p_spTree = target

        inSlide_id = len(p_spTree)
        print str(inSlide_id) +' indSlide'



        #make the picture tree structure#
        p_pic       = Element('{http://schemas.openxmlformats.org/presentationml/2006/main}pic')

        p_nvPicPr   = ET.SubElement(p_pic,'{http://schemas.openxmlformats.org/presentationml/2006/main}nvPicPr') #in p_pic
        p_cNvPr     = ET.SubElement(p_nvPicPr,'{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr', {'id':str(inSlide_id),'name':'Content Placeholder X'}) #in p_nvPicPr
        p_cNvPicPr  = ET.SubElement(p_nvPicPr,'{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPicPr') #in p_nvPicPr
        a_picLocks  = ET.SubElement(p_cNvPicPr,'{http://schemas.openxmlformats.org/drawingml/2006/main}picLocks',{'noGrp':"1", 'noChangeAspect':"1"})
        p_nvPr      = ET.SubElement(p_nvPicPr,'{http://schemas.openxmlformats.org/presentationml/2006/main}nvPr')
        p_ph        = ET.SubElement(p_nvPr,'{http://schemas.openxmlformats.org/presentationml/2006/main}ph',{'sz':"half",'idx':"1"})

        p_blipFill  = ET.SubElement(p_pic,'{http://schemas.openxmlformats.org/presentationml/2006/main}blipFill')
        a_blip      = ET.SubElement(p_blipFill,'{http://schemas.openxmlformats.org/drawingml/2006/main}blip',{'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link':relationship_id})
        a_stretch   = ET.SubElement(p_blipFill,'{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
        a_fillRect  = ET.SubElement(a_stretch,'{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect')

        p_spPr      = ET.SubElement(p_pic,'{http://schemas.openxmlformats.org/presentationml/2006/main}spPr')
        a_xfrm      = ET.SubElement(p_spPr,'{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
        a_off       = ET.SubElement(a_xfrm,'{http://schemas.openxmlformats.org/drawingml/2006/main}off',{'x':"952500",'y':"2910681"})
        a_ext       = ET.SubElement(a_xfrm,'{http://schemas.openxmlformats.org/drawingml/2006/main}ext',{'cx':"3048000",'cy':"1905000"})


        p_spTree.insert(len(p_spTree)-1,p_pic)



        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/ppt/slides/slide" + file_nr + ".xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/ppt/slides/slide" + file_nr + ".xml", "a") as text_file:
            text_file.write(xmlstr)


        self.zipdir(self.work_dir, self.outputfile)  
        return True         


    ET.register_namespace('wpc', "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
    ET.register_namespace('mc', "http://schemas.openxmlformats.org/markup-compatibility/2006")
    ET.register_namespace('o', "urn:schemas-microsoft-com:office:office")
    ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ET.register_namespace('m', "http://schemas.openxmlformats.org/officeDocument/2006/math")
    ET.register_namespace('v', "urn:schemas-microsoft-com:vml")
    ET.register_namespace('wp14', "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
    ET.register_namespace('wp', "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
    ET.register_namespace('w10', "urn:schemas-microsoft-com:office:word")
    ET.register_namespace('w', "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    ET.register_namespace('w14', "http://schemas.microsoft.com/office/word/2010/wordml")
    ET.register_namespace('wpg', "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
    ET.register_namespace('wpi', "http://schemas.microsoft.com/office/word/2010/wordprocessingInk")
    ET.register_namespace('wne', "http://schemas.microsoft.com/office/word/2006/wordml")
    ET.register_namespace('wps', "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
    ET.register_namespace('Ignorable', "w14 wp14")
    ET.register_namespace('mc:Ignorable', "w14 wp14")
    ET.register_namespace('Relationships', "http://schemas.openxmlformats.org/package/2006/relationships")
    ET.register_namespace('a', "http://schemas.openxmlformats.org/drawingml/2006/main")
    ET.register_namespace('pic', "http://schemas.openxmlformats.org/drawingml/2006/picture")
    ET.register_namespace('p',"http://schemas.openxmlformats.org/presentationml/2006/main")
    ET.register_namespace('p14',"http://schemas.microsoft.com/office/powerpoint/2010/main")




