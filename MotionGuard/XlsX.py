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


class XlsX(Document.Document):
    """M$ Office Spreadsheet"""
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

    def check(self,looking,textfile):
            datafile = file(textfile)
            found = False #this isn't really necessary 
            for line in datafile:
                if looking in line:
                    #found = True #not necessary 
                    return True
            return False #because you finished the search without finding anything

    def inject_bait(self):
        partname = 'drawingX.xml'
        partnameID =1;
        found_it = False
        relationship_id = 'rId1'
        relationship_id_cntr =1
        draw_found = True
        picture_name = ''   ##in drawing1.xml, individual and independent in this file
        picture_name_cntr = 1 


        

        #########################################
        #######./[Content_Types].xml##############
        #########################################
        ##make a valid id nr
        while(not found_it): #if good solution founded
            if (self.check('drawing'+str(partnameID)+'.xml',self.work_dir + '/[Content_Types].xml')):
                partnameID +=1
            else:
                found_it = True

    
        document_tree = ET.parse(self.work_dir+ "/[Content_Types].xml" )
        root = document_tree.getroot()

        partname = 'drawing' +str(partnameID)+ '.xml'
        root.append(Element('Override',{'PartName':'/xl/drawings/'+partname,'ContentType':'application/vnd.openxmlformats-officedocument.drawing+xml'}))


        ###saving the file##############################
        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/[Content_Types].xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/[Content_Types].xml", "a") as text_file:
            text_file.write(xmlstr)

        self.replace(self.work_dir+ "/[Content_Types].xml",'ns0:','') ##elementtree namespace dont work optimally
        self.replace(self.work_dir+ "/[Content_Types].xml",':ns0','')
        found_it=False



        ####################################################
        ############end of ./[Content_Types].xml###########
        ##################################################




        ######################################################################
        ##make worksheet/_rels dics############################################
        ##################################################################
        
        if(not os.path.isdir(self.work_dir + '/xl/worksheets')):
            os.makedirs(self.work_dir +'/xl/worksheets')

        if(not os.path.isdir(self.work_dir +'/xl/worksheets/_rels')):
            os.makedirs(self.work_dir +'/xl/worksheets/_rels')


        if(not os.path.exists(self.work_dir +'/xl/worksheets/_rels/sheet1.xml.rels')): #make the file structure
            with open(self.work_dir +'/xl/worksheets/_rels/sheet1.xml.rels', 'w+') as f:
                 f.close
            root = Element('Relationships',{'xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'})
            rel = ET.SubElement(root,'Relationship',{'Id':str(relationship_id),'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing','Target':self.work_dir + '/xl/drawings/drawing1.xml'})
        else:
            root = Element('Relationships',{'xmlns':'http://schemas.openxmlformats.org/package/2006/relationships'}) #if there are the filestructure yet
            while(not found_it): #if good solution founded
                if (self.check('rId'+str(relationship_id_cntr),self.work_dir + '/xl/worksheets/_rels/sheet1.xml.rels')): #make a valid id for the picture
                    relationship_id_cntr +=1
                else:
                    found_it = True
                    relationship_id = 'rId' +str(relationship_id_cntr)
                    rel = ET.SubElement(root,'Relationship',{'Id':str(relationship_id),'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing','Target':self.work_dir + '/xl/drawings/drawing1.xml'})


        xmlstr = ET.tostring(root)
        with open(self.work_dir +"/xl/worksheets/_rels/sheet1.xml.rels", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/xl/worksheets/_rels/sheet1.xml.rels", "a") as text_file:
            text_file.write(xmlstr)
        found_it = False
        #########END of /xl/worksheets/_rels/sheet1.xml.rels ###########################
        #################################################################################
        
        ##################xl/worksheets/sheet1.xml######################################
        ##############################################################################
        document_tree = ET.parse(self.work_dir +"/xl/worksheets/sheet1.xml" )
        root = document_tree.getroot()
        drawing = ET.SubElement(root,'drawing',{'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id':relationship_id})

        xmlstr = ET.tostring(root)
        with open(self.work_dir +"/xl/worksheets/sheet1.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/xl/worksheets/sheet1.xml", "a") as text_file:
            text_file.write(xmlstr)

        self.replace(self.work_dir +"/xl/worksheets/sheet1.xml",'ns0:','') ##elementtree namespace dont work optimally
        self.replace(self.work_dir +"/xl/worksheets/sheet1.xml",':ns0','')
        found_it=False

        ###END of sheet1.xml##################################################################x

        ##################/xl/DRAWINGS directory###########################
        draw_found = True
        if(not os.path.isdir(self.work_dir + '/xl/drawings')):
            os.makedirs(self.work_dir + '/xl/drawings')
            draw_found = False

        if(not os.path.exists(self.work_dir + '/xl/drawings/drawing1.xml')):
            open(self.work_dir + '/xl/drawings/drawing1.xml','a').close
            draw_found = False

        if(not os.path.isdir(self.work_dir + '/xl/drawings/_rels')):
            os.makedirs(self.work_dir + '/xl/drawings/_rels')
            draw_found = False

        if(not os.path.exists(self.work_dir + '/xl/drawings/_rels/drawing1.xml.rels')):
            open(self.work_dir + '/xl/drawings/_rels/drawing1.xml.rels','a').close
            draw_found = False

              
        root = None
        if(draw_found):
            document_tree = ET.parse(self.work_dir + "/xl/drawings/_rels/drawing1.xml.rels")
            root = document_tree.getroot()
        if (root == None):
            root = Element('{http://schemas.openxmlformats.org/package/2006/relationships}Relationships')
        ET.SubElement(root,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',{'Id':relationship_id,'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image','Target':self.bait,'TargetMode':'External'})

        
        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/xl/drawings/_rels/drawing1.xml.rels", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir +"/xl/drawings/_rels/drawing1.xml.rels", "a") as text_file:
            text_file.write(xmlstr)

        self.replace(self.work_dir +"/xl/drawings/_rels/drawing1.xml.rels",'Relationships:','')
        self.replace(self.work_dir +"/xl/drawings/_rels/drawing1.xml.rels",'xmlns:Relationships','xmlns')
          
        ###xl\drawings\drawing1.xlsx################
        found_it=False
        while(not found_it): #if good solution founded
            if (self.check('Picture '+str(picture_name_cntr),self.work_dir + "/xl/drawings/drawing1.xml")): #make a valid id for the picture
                picutre_name_cntr +=1 #if there are this id, then try another
            else:
                found_it = True #if these id is not exist yet
                picture_name = 'Picture ' +str(picture_name_cntr)
                rel = ET.SubElement(root,'Relationship',{'Id':str(picture_name),'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing','Target':self.work_dir + "/xl/drawings/drawing1.xml"})



        root = None
        if(not draw_found):
            root = Element('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}wsDr')
        else:
            document_tree = ET.parse(self.work_dir + "/xl/drawings/drawing1.xml")
            root = document_tree.getroot()

        xdr_twoCell = Element('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor',{'editAs':'oneCell'})
        xdr_from = ET.SubElement(xdr_twoCell,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from')

        xdr_col = ET.SubElement(xdr_from,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col').text='0'
        #xdr_col.text = '0'
        xdr_colOff = ET.SubElement(xdr_from,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}colOff').text='0'
        #xdr_colOff.text ='0'
        xdr_row =   ET.SubElement(xdr_from,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row').text='0'
        xdr_rowOff =   ET.SubElement(xdr_from,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}rowOff').text='0'


        xdr_to =ET.SubElement(xdr_twoCell,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}to')
        xdr_col = ET.SubElement(xdr_to,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col').text='2'
        xdr_colOff = ET.SubElement(xdr_to,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}colOff').text='2'
        xdr_row =   ET.SubElement(xdr_to,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row').text='2'
        xdr_rowOff =   ET.SubElement(xdr_to,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}rowOff').text='2' ##place of the picture

        xdr_pic =ET.SubElement(xdr_twoCell,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic')
        xdr_nvPicPr = ET.SubElement(xdr_pic,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}nvPicPr')
        xdr_cNvPr =  ET.SubElement(xdr_nvPicPr,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPr',{'id':'111111','name':picture_name})
        xdr_cNvPicPr = ET.SubElement(xdr_nvPicPr,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cNvPicPr')
        a_picLocks = ET.SubElement(xdr_cNvPicPr,'{http://schemas.openxmlformats.org/drawingml/2006/main}picLocks',{'noChangeAspect':'1'})

        xdr_blipFill= ET.SubElement(xdr_pic,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}blipFill')
        a_blip  = ET.SubElement(xdr_blipFill,'{http://schemas.openxmlformats.org/drawingml/2006/main}blip',{'{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link':relationship_id})
        a_stretch   =ET.SubElement(xdr_blipFill,'{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
        a_fillRect  =ET.SubElement(a_stretch,'{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect')

        xdr_spPr    =ET.SubElement(xdr_pic,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}spPr')
        a_xfrm      =ET.SubElement(xdr_spPr,'{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
        a_off       =ET.SubElement(a_xfrm,'{http://schemas.openxmlformats.org/drawingml/2006/main}off',{'x':'0','y':'0'})
        a_ext       =ET.SubElement(a_xfrm,'{http://schemas.openxmlformats.org/drawingml/2006/main}ext',{'cx':'1695450','cy':'1695450'}) ##size of the picture
        a_prstGeom  =ET.SubElement(xdr_spPr,'{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom',{'prst':'rect'}) 
        a_avLst     =ET.SubElement(a_prstGeom,'{http://schemas.openxmlformats.org/drawingml/2006/main}avLst') 

        xdr_clientData = ET.SubElement(xdr_twoCell,'{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}clientData')

        root.append(xdr_twoCell)

        xmlstr = ET.tostring(root)
        #write xml header
        with open(self.work_dir + "/xl/drawings/drawing1.xml", "w") as text_file:
            text_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')

        #write xml body
        with open(self.work_dir + "/xl/drawings/drawing1.xml", "a") as text_file:
            text_file.write(xmlstr)

        ####end xl\drawings\drawing1.xlsx

        #End of /xl/drawings files and dirs#####################
        ########################################################
        

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
    ET.register_namespace('x14ac','http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')  
    ET.register_namespace('xdr','http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')



