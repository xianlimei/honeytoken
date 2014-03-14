import xml.etree.ElementTree as ET
from xml.etree.ElementTree import tostring, Element
import zipfile
import shutil
import os 
import sys
from optparse import OptionParser
from contextlib import closing
import sqlite3
import datetime
import md5

import Ods
import Odp
import PptX
import XlsX
import Odt
import DocX

from tempfile import mkstemp

used_types = ('ods','odp','xlsx','pptx','docx','odt')
time = ''
alert_trigger = 'aI38tX'
ip_address ='10.105.0.49'
server_path = '/var/www/'

def input_validate(inp, outp, wd,ba, db):
    try:
        print 'Validate:________________'
        print 'input: '+inp
        print 'out '+outp
        print 'working: '+wd 
        print 'bait: '+ba
        print 'database' + db
        print'_________________________'
        if(inp == '' or db == ''):  #if input string is empty
            print 'Error! Missing necessary argument. Missing input file path or database path.'
            sys.exit(1)
        else:

            if(inp != ''):# if input file path is not empty
                exit_trigger = True;
                for type in used_types:
                    if (inp.endswith(type)):
                        exit_trigger = False
                if(exit_trigger == True):
                     print 'Error! Wrong filetype'
                     sys.exit(1)
    

                inp = inp.replace('\\','/') #if inp not empty
            if(outp !=''): #if output exists, then must be the same format
                outp = outp.replace('\\','/')

            if(db!=''):
                db = db.replace('\\','/')

            if(outp !='' and inp !=''): #if we have input and output file, both must be the same type
              #  if(not((inp.endswith('.docx') and outp.endswith('.docx')) or (inp.endswith('odt') and outp.endswith('.odt')))):
               if(not((inp.endswith('.xlsx') and outp.endswith('.xlsx')) or (inp.endswith('ods') and outp.endswith('.ods'))
                    or (inp.endswith('.docx') and outp.endswith('.docx')) or (inp.endswith('odt') and outp.endswith('.odt'))
                    or (inp.endswith('pptx') and outp.endswith('.pptx')) or (inp.endswith('odp') and outp.endswith('.odp')))) :
                        print 'Error! The input and output file are different format!'
                        sys.exit(1)


            if(outp =='' and inp != '' ): #if output is missing
                if(inp.endswith('.docx')):
                   outp = inp[:-5] +'_injected.docx'   #then we create an output file
                if(inp.endswith('.odt')):
                    outp = inp[:-4] +'_injected.odt'
                if(inp.endswith('.xlsx')):
                    outp = inp[:-5] +'_injected.xlsx'   #then we create an output file
                if(inp.endswith('.ods')):
                    outp = inp[:-4] +'_injected.ods'
                if(inp.endswith('.pptx')):
                    outp = inp[:-5] +'_injected.pptx'   #then we create an output file
                if(inp.endswith('.odp')):
                    outp = inp[:-4] +'_injected.odp'


            if(wd !='' and (wd.endswith('/') or wd.endswith('\\')) ): #if we gave a working dir path and it has an end with / then we cut it off
                wd = wd[:-1]
            else:
                wd = wd.replace('\\','/')
                wd = inp.rsplit('/', 1)[0]
                ######################################problema ha egy mappaban vagyunk az inputtal, azt kulonn kezelkni kell ha ures a workdir es nincs / az inputban
                ###############################################
                ########################################

            wd =  wd    +"/workdir" 
            time = datetime.datetime.now().strftime("%y-%m-%d-%H-%M") #the time when we inject the file into the database
            hash_print = md5.new(''+inp + outp + ba + time).hexdigest() #make hash from the datas
        return (inp, outp, wd, ba,db,hash_print)
    except IOError as io:
        print "I/O error({0}): {1} in input_validate".format(io.errno, io.strerror)
    except:
        print "Unexpected error:", sys.exc_info()[0]


def updateDatabase(inp,outp,wd,bait,db,table_name,hash_print):
    """make the baitfile and updated the database with these information"""
    try:
       # hashed_bait = "L:\\" + str(hash_print) +str(alert_trigger) + ".jpg"
        hashed_bait = server_path + str(hash_print) +str(alert_trigger) + ".jpg" #to copy into the right folder
        #############ALERT make a linux compatible directory name
        conn = None

        if(not (os.path.exists(bait) and os.path.exists(server_path) )):  ##os.path.exists(hashed_bait)
            print '-----------------GhostDecoy----------------'
            print 'old or new path of the picture doesnt exists'
        else:
            print 'copy correct to: '  + hashed_bait
            shutil.copyfile(bait,hashed_bait)
 
  

        conn = sqlite3.connect(db)

        ###create table if not exists
        sql = 'create table if not exists ' + table_name +'(input text NOT NULL, output text NOT NULL, bait text NOT NULL,time text NOT NULL, hash text primary key NOT NULL )'
        curs = conn.cursor()
        curs.execute(sql)
        conn.commit()
        ##################

        curs = conn.cursor()
        # curs.execute("SELECT * FROM " + table_name + " WHERE hash = ?", (hash,))
        curs.execute("SELECT * FROM " + table_name + " WHERE hash = '" +hash_print +"';")
        data=curs.fetchone()
        if data is None:
            print('There is no component hashed %s'%hash_print)
            curs.execute("INSERT INTO " +table_name + " (input,output,bait,time,hash) VALUES('" +inp + "','" +outp+"','" +bait+"','" +time+ "','"+hash_print+ "');")
            conn.commit()
            hashed_bait = "http://" + ip_address +'/'+ str(hash_print) +str(alert_trigger) + ".jpg" #hashed_bait to inject, we nedd the ip address
            bait = hashed_bait

            conn.close()
            return True
        else:
            print('Component %s found with datas %s %s %s %s'%(hash_print,data[0],data[1],data[2],data[3]))
            conn.close()
            return False 
    except IOError as io:
        print "I/O error({0}): {1} in updateDatabase".format(io.errno, io.strerror)
    except:
        print "Unexpected error in updateDatabase:", sys.exc_info()[0]



def main(argv):
    inputfile = ''
    outputfile = ''
    workingdir = ''
    bait = ''
    database  = ''
    hash_print =''
    can_work = False;


    try:
        parser = OptionParser()
        parser.add_option(
            '-i', '--input',
            dest = 'inputfile',
            help = 'the path of the input file',
            metavar = 'INPUT_FILE'
        )

        parser.add_option(
            '-o', '--output',
            dest = 'outputfile',
            help = 'the path of the output file',
            metavar = 'OUTPUT_FILE'
        )

        parser.add_option(
            '-w', '--workdir',
            dest = 'workingdir',
            help = 'the path of the work directory',
            metavar = 'WORK_DIRECTORY'
        )

        parser.add_option(
            '-b', '--bait',
            dest = 'bait',
            help = 'the path of the bait picture',
            metavar = 'BAIT_PIC'
        )

        parser.add_option(
            '-d', '--database',
            dest = 'database',
            help = 'the path of the sqlite database',
            metavar = 'DATABASE'
        )



        parser.set_defaults(inputfile = '')
        parser.set_defaults(outputfile = '')
        parser.set_defaults(workingdir = '')
        parser.set_defaults(bait = '')
        parser.set_defaults(database = '')
        opts, args = parser.parse_args()

        inputfile = opts.inputfile
        outputfile = opts.outputfile
        workingdir = opts.workingdir
        bait = opts.bait
        database = opts.database



        #validate and make the incoming arguments
        inputfile, outputfile, workingdir, bait ,database,hash_print = input_validate(inputfile,outputfile,workingdir,bait,database) #variables after validate

        hashed_bait = "http://" + ip_address+'/' + str(hash_print) +str(alert_trigger) + ".jpg"
        print("Python Injecter")

        if(bait==""):
            bait='L:\wallpaper253534.jpg'

        can_work=False

        # build a tree structure
        if(inputfile.endswith('.ods')):
            choosen_document = Ods.Ods(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)
        if(inputfile.endswith('.odt')):
            choosen_document = Odt.Odt(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)
        if(inputfile.endswith('.odp')):
            choosen_document = Odp.Odp(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)
        if(inputfile.endswith('.docx')):
            choosen_document = DocX.DocX(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)
        if(inputfile.endswith('.pptx')):
            choosen_document = PptX.PptX(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)
        if(inputfile.endswith('.xlsx')):
            choosen_document = XlsX.XlsX(inputfile,outputfile, workingdir,hashed_bait)
            can_work = choosen_document.extract_file(outputfile, inputfile,workingdir)

    
        if(can_work == True):
            print("Extract done, inject come")
            raw_input("Press any button")
            ok = False
            ok = choosen_document.inject_bait()
            if (ok):
                print 'OK\n'
                print updateDatabase(inputfile,outputfile,workingdir,bait,database,"injected_documents",hash_print)
                print '\n'

            raw_input("Press any button to delete ")
            shutil.rmtree(workingdir)
            print("Delete OK")


        else:
            print "Nothing to do here"
        raw_input("Press any button to exit. ")

    except IOError as io:
        print "I/O error({0}): {1} in MAIN".format(io.errno, io.strerror)
    except:
        print "Unexpected error in main:", sys.exc_info()[0]


if __name__ == "__main__":
   main(sys.argv[1:])
  # raw_input()