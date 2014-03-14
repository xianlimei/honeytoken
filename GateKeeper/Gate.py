import os
import sqlite3
import sys



def check_consistency(hash,db):
    conn = None
    try:
        conn = sqlite3.connect(db)
        cur = conn.cursor()    
        cur.execute("SELECT * FROM injected_documents WHERE hash='" + hash + "';")
        data=cur.fetchone()
        if data is None:
            print "Not Found"
        else:
            print "Founded!"
            print data
            #what are we doing with the result
            ###################################
            ####################################

    except sqlite3.Error, e:
        print "Error %s:" % e.args[0]
        sys.exit(1)
    
    finally:  
       if conn:
          conn.close()

def main(argv):
    start_from = 0
    alert_trigger = 'aI38tX'
    db = '/home/toke/tst.db'          #generated database file with all of the injected documents
    alert_file = 'alert.txt'    #in this file are all of the alerts, with end of 'aI38tX'
    marker = 'marker.txt'       #number of used rows, we dont want to read it once again

    #the file format: full_path + hash as file_name + alert_trigger(aI38tX) + .extension
    #like int win L:\833e3f998ae1f729cd820db52b8568ccaI38tX.jpg

    try:
        start_from = 0
        if(not os.path.exists(marker)):
             with open(marker,'w') as mrk:
                mrk.write('0')
                mrk.close()
        with open(marker,'r') as mrk:
            start_from = int( mrk.readline())
            mrk.close()
    except:
        pass


    with open(alert_file, 'r') as file:
        for i, line in enumerate(file):
            if i <= start_from: continue
            else:
                columns = line.strip().split(' ')
                start_from =i
                for col in columns:
                    if(col.find(alert_trigger)>0):    #return -1 if didnt find it
                        trig1 = col.split('/')  #cut the full path
                        trig2 = trig1[len(trig1)-1]
                        trigger = os.path.splitext(trig2)[0] #trigger is only the hash value whitout full path or extension, cut the extension 
                        if trigger.endswith(alert_trigger):
                            trigger = trigger[:-(len(alert_trigger))] # trigger without  alert_trigger = 'aI38tX', cut the trigger
                        check_consistency(trigger,db)


    try:
        with open(marker,'w') as mrk:
            start_from = str(start_from)
            mrk.write(start_from)
            mrk.close()
    except:
        pass

    raw_input()

if __name__ == "__main__":
   main(sys.argv[1:])