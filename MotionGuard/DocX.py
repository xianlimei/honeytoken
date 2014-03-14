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

class DocX(Document.Document):
   inputfile = ''
   outputfile = ''
   work_dir = ''
   bait = ''

   def __init__(self, input, output, work, bait):
       self.inputfile = input
       self.outputfile = output
       self.work_dir = work
       self.bait = bait
       print self.inputfile
       print self.outputfile
       print "sadasd"

   def inject_bait(self):
            relationship_id = "rIdf"    #ID for the relationship in /word/_rels/document.xml.rels
            found =0
            can_write=0 #when we are through the document tag, we can write the tree into the new file
            rsIdr_num = '001F25BA'

            document_header = '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office "'
            document_header = document_header + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml "' 
            document_header = document_header +' xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word "' 
            document_header = document_header+ ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup "' 
            document_header = document_header+ ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14 "'
            document_header = document_header + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" >'
        
            #make the /word/_rels/document.xml.rels
            document_tree = ET.parse(self.work_dir+"/word/_rels/document.xml.rels")
            root = document_tree.getroot()
            element = root.find('Relationships')
            root.append(Element('Relationship',{'Id':relationship_id,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image','Target':self.bait,'TargetMode':'External'}))
        

            xmlstr = ET.tostring(root)
        
            #write xml header
            with open(self.work_dir +"/word/_rels/document.xml.rels","w") as text_file:
              text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

            #write xml body
            with open(self.work_dir +"/word/_rels/document.xml.rels", "a") as text_file:
             text_file.write(xmlstr)
            self.replace(self.work_dir +"/word/_rels/document.xml.rels",'Relationships:','')
            self.replace(self.work_dir +"/word/_rels/document.xml.rels",'xmlns:Relationships','xmlns')
        

            ##########Write the document file####################
            document_tree = ET.parse(self.work_dir+"/word/document.xml")
            root = document_tree.getroot()
            body= root.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')

            body.append(Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p',{'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR':rsIdr_num}))
            for p in body.findall( '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p' ):
               if '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR' in p.attrib:       ##python on linux cant use the attribute,it's dropping a KeyError error, because not all of the 'p' has an attribute
                    if(p.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR']==rsIdr_num):
                        wp =p

            wp.append(Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'))
            wr = wp.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        
            wr.append(Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'))
            draw = wr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
        
            draw.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline'))
            inline = draw.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')

            inline.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent',{'cx':'5372643','cy':'3019425'}))
            inline.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}effectExtent',{'l':'0','t':'0','r':'0','b':'0'}))
            inline.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr',{'id':'0','name':'inj_pic.jpg'}))
            inline.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}cNvGraphicFramePr'))

            wpFramePr = inline.find('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}cNvGraphicFramePr')
            wpFramePr.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicFrameLocks',{'noChangeAspect':'1'}))

            inline.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}graphic'))
            agraph = inline.find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
          
            agraph.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData',{'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'}))
            gData = agraph.find('{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
            gData.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}pic'))

            pic = gData.find('{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')
            pic.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr'))

            nvPic = pic.find('{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr')
            nvPic.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr',{'id':'0','name':'inj_pic.jpg'}))
            nvPic.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr'))

            pic.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill'))
            blipFill = pic.find('{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
            blipFill.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}blip',{'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link':relationship_id}))
            blipFill.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}stretch'))
            stretch = blipFill.find('{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
            stretch.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect'))

            pic.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr'))
            spPr = pic.find('{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr')
            spPr.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm'))
            xfrm = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
            xfrm.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}off',{'x':'0','y':'0'}))
            xfrm.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}ext',{'cx':'5372643','cy':'3019425'}))

            spPr.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom',{'prst':'rect'}))
            prst = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
            prst.append(Element('{http://schemas.openxmlformats.org/drawingml/2006/main}avLst'))
        
        

       

        #    with open(work_dir +"/word/document.xml","a") as text_file:
         #       text_file.write(document_header) 
        

            xmlstr = ET.tostring(root)
        
            with open(self.work_dir +"/word/document.xml", "w") as text_file:
             text_file.write(xmlstr)

            #del the document element and replace with the right document element
            doc_file =open(self.work_dir+"/word/document.xml",'r')
            tmp_file =open(self.work_dir+"/word/tmp.xml",'w')
            for char in doc_file.read():
                if(can_write==1):
                    tmp_file.write(char)
                if(char=='>'):
                    can_write=1
            doc_file.close()
            tmp_file.close()

            with open(self.work_dir +"/word/document.xml","w") as text_file:
              text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')     
              text_file.write(document_header) 
        
            tmp_file = open(self.work_dir+"/word/tmp.xml",'r')
            full_content = tmp_file.read()
            tmp_file.close()
            os.remove(self.work_dir+"/word/tmp.xml") #delete the temporary file


            with open(self.work_dir +"/word/document.xml","a") as text_file:
                text_file.write(full_content)
            print("The decoy injection is done!")
            raw_input()
        
            self.zipdir(self.work_dir, self.outputfile)
        




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
