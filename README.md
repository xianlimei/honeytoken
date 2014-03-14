honeytoken
==========

A simple honeytoken injection script written in Python. It can be inject into docx, pptx, xslx, odt, odp, ods documents.


F.1 UserGuidetoMotionGuard.py
On many systems are installed the following modules, but on some systems mainly on Windows these modules are often missing.
To run the script successful the machines need the ElementTree xml parser and the Sqlite3 module for the databases. 
The script are written in python so we must install the python environment too.
I recommended to run the program with root privileges because it make and delete some directory and sometimes it can be done with user privileges.
Operation: running with the MotionGuard.py ﬁle 
• -h: help tab 
• -i: the path of the input ﬁle 
• -o: the path of the output ﬁle 
• -b: the path of the bait picture 
• -w: path of the working directory 
• -d: path of the database where the datas will be saved 
• -a: IP address of the server machine 
• -s: network directory of the sever, where the images can be downloaded from 

pathComment: the -i, -b, -d, -a is required 
If output ﬁle is blank than it will be the inputﬁle + the ’_injected’string.If the working directory areemptythanitwillbethepathoftheinputﬁle+/workdirdirectory.Thenetworkdirofthe server is default /var/www which is the default Apache server directory.
The program sometimes needed ’enter’ during the injection procedure. 
This function stayed at the code for demonstration purposes.

F.2 UserGuidetoGate.py
If we are runned the MotionGuard.py and the log watcher code yet, then all of the ﬁles are ready to run the Gate.py script.
Log watcher code: 
tail -f /var/log/apache2/access.log | egrep –line-buffered aI38tX | tee -a /home/alert tail -f (path of access.log) tee -a (path of the alert ﬁle, we choose this ﬁle) Gate.py operation: running with the Gate.py ﬁle
• -h: help tab • -d:  of the database what is made by the MotionGuard.py earlier • -a: path of the alert ﬁle which will be ﬁlled with the alerts by the log watcher script • -m:pathofthemarkerﬁle.Itwillcontainonlyonenumberwhichrepresentstherowweare used the last
