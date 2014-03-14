import xml.etree.ElementTree as ET
from xml.etree.ElementTree import tostring, Element
import zipfile
import shutil
import os 
import sys, getopt
from contextlib import closing
import sys, getopt
from tempfile import mkstemp


class Document(object):
    """It is the parent class of the concret document types"""
    inputfile = ''
    outputfile = ''
    workingdir = ''
    bait = ''
    can_work = False;

    def __init__(self, input, output, work, bait):
        #Construct
        self.inputfile = input
        self.outputfile = output
        self.workingdir = work
        self.bait = bait
        self.can_work = False;


    def replace(self,file_path, pattern, subst):
        #used only with M$ filetypes
        #Create temp file
        fh, abs_path = mkstemp()
        new_file = open(abs_path,'w')
        old_file = open(file_path)
        for line in old_file:
            new_file.write(line.replace(pattern, subst))
        #close temp file
        new_file.close()
        os.close(fh)
        old_file.close()
        #Remove original file
        os.remove(file_path)
        #Move new file
        shutil.move(abs_path, file_path)


    def zipdir(self, basedir, archivename):
        #zip
        assert os.path.isdir(basedir)
        with closing(zipfile.ZipFile(archivename, "w", zipfile.ZIP_DEFLATED)) as z:
            for root, dirs, files in os.walk(basedir):
                #NOTE: ignore empty directories
                for fn in files:
                    absfn = os.path.join(root, fn)
                    zfn = absfn[len(basedir)+len(os.sep):] #XXX: relative path
                    z.write(absfn, zfn)

    def extract_file(self,output_file, input_file,work_dir):
        #unzip
        if(output_file and input_file and work_dir):
                #extract the files
            zip = zipfile.ZipFile(r''+ input_file+'') #path of the odt file
            zip.extractall(r''+work_dir+'')   #target directory name for the unzipped file
            print("File unzipped!")

            print(input_file +'\n')
            print(output_file + '\n')
            return True  
        else:
            return False   

    def inject_bait(self):
        """override"""