#Coded for Python 3.4.3
from tkinter import *
import re, linecache, os
import xlwt, xlrd

#For practice results, use SP-20150725-003
#   It has a good amount of everything to parse through

#Created by Myles Morrone
#Ver(2.0) 12:05 8/3/2015            
#The xlwt module is available at: https://pypi.python.org/pypi/xlwt
#The xlrd module is available at: https://pypi.python.org/pypi/xlrd

#Current build info:
#1. Automatic connection to the OneWiki does not work. (Cannot get mwclient to work locally)
#2. Lines with unreadable characters WILL BE IGNORED, be sure to review these manually!
#3. Additional files and MD5s, found in the File Analysis section, will NOT be automatically related.

#Instructions
#1. Create a new text file named "ToBeParsed.txt" in same folder as this file
#2. Go to desired Wiki page and click the "Edit" tab
#3. Copy and paste ALL the information from the edit box into your text file (Ctrl+A, Ctrl+C, Ctrl+V)
#4. Run this program (F5 in IDLE or double click if saved as a .pyw file)
#5. Input Ticket Number, SP ID, comments, and finally click "Start"

#JtR will parse data, create necessary reference text files, input into excel, open excel file for viewing
#   and clean up reference files. Once reviewing has been finished, save excel file in .csv format. To
#   begin another CRIT report, you must restart JtR (completely flushing memory).

#GUI creation using Tkinter
root = Tk()
root.title("Jack the Ripper")
root.configure(bg="black")
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
JtR = "JtR.ico"

#GUI Configuration Variables
GUI_W = 370 #GUI width -- 370 default, recommended to not change
GUI_H = 190 #GUI height -- 190 default, recommended to not change
Screen = 2 #1=Leftmost, 2=Middle, 3=Rightmost]
GUI_P = 1 #1=Top left; 2=Top right; 3=Bottom left; 4=Bottom right


GUI_X = sw*0.05 #GUI X Position on screen
GUI_Y = sh*0.05 #GUI Y Position on screen
if GUI_P < 1 or GUI_P > 4:
    GUI_P = 1
if Screen < 1 or Screen > 3:
    Screen = 1
if GUI_P % 2 != 1: #Checks for GUI_P being even, left or right side of screen
    GUI_X = (sw*0.95)-GUI_W
if GUI_P > 2: #Checks for GUI_P being greater than 2, top or bottom of screen
    GUI_Y = (sh*0.90)-GUI_H
try:
    root.wm_iconbitmap(JtR)
except:
    pass

#                             (Width, height, Xpos, Ypos)
root.geometry("%dx%d+%d+%d" % (GUI_W, GUI_H, ((GUI_X)+(sw*(Screen-1))), GUI_Y))

#GUI Vars
RT = StringVar()
#RT.set("123456") #For demo purposes
SP = StringVar()
#SP.set("SP-19901203-007") #For demo purposes

#Global Vars
TBP = "Files/ToBeParsed.txt"
Results = "Files/Results.txt"
LineAgg = "Files/LineAgg.txt"

extList = [   ".doc:",  ".docx:",   ".log:",   ".msg:",   ".odt:",
              ".rtf:",   ".tex:",   ".txt:",   ".csv:",   ".dat:",
              ".pps:",   ".ppt:",   ".vcf:",   ".xml:",   ".bmp:",
              ".gif:",   ".jpg:",   ".png:",   ".tif:",   ".pct:",
              ".pdf:",   ".xlr:",   ".xls:",  ".xlsx:",    ".db:",
              ".dbf:",   ".mdb:",   ".sql:",   ".exe:",   ".jar:",
              ".pif:",    ".vb:",   ".vbs:",   ".asp:",   ".cfm:",
              ".css:",   ".htm:",  ".html:",    ".js:",   ".jsp:",
              ".php:", ".xhtml:",   ".cfg:",   ".ini:",    ".7z:",
              ".deb:",    ".gz:",   ".pkg:",   ".rar:",   ".rpm:",
           ".tar.gz:",   ".zip:",  ".zipx:",".exifdata:"] #extensions to split file names and MD5s

domList = [    ".com/",    ".org/",    ".net/",    ".int/",    ".edu/",
               ".gov/",    ".mil/",   ".arpa/"] #top-level domain list

def JtR(): #Info parser - Jack the Ripper
    print("-----JtR START-----")
    maxL = sum(1 for line in open(TBP)) #Counts total lines in ToBeParsed.txt
    print(str(maxL) + " lines to be parsed by JtR")
    L = 1 #Current line, for iterating
    LP = 0 #Lines printed, mainly for debugging purposes
    P = False #Print boolean
    XLtype = "" #Input for Type column in excel file
    XLrole = "" #Input for Role column in excel file
    with open(Results, "w") as fileW:
        while L < maxL+1:
            try:
                line = linecache.getline(TBP, L) #Pull line #"L" from cached source
                line = line[:-1] #Remove newline character (\n) from end of line
                if re.search(r"�", line): #Obliterate these annoying little characters
                    line = "" #!!!!NOTICE!!!! These lines are ANNIHILATED, manually input them if you want them
                if re.search(r"Notable", line): #Has parsed too far, halt parsing
                    break #Do not pass GO, do not collect $200
                if re.match(r"X-Mailer", line): #Section Headers, automatically changes inputs as per type
                    XLtype = "Email Header - X-Mailer"
                    XLrole = ""
                if re.match(r"Sender domain", line):
                    XLtype = "URI - Domain Name"
                    XLrole = ""
                if re.match(r"Sender IP", line):
                    XLtype = "Address - ipv4-addr"
                    XLrole = "Sender_IP"
                if re.match(r"Sender mail", line):
                    XLtype = "Address - e-mail"
                    XLrole = "Sender_Address"
                if re.match(r"Subject", line):
                    XLtype = "Email Header - Subject"
                    XLrole = ""
                if re.match(r"Attachment names", line):
                    XLtype = "Hash - MD5"
                    XLrole = "Attachment"
                if re.match(r"Message body links", line):
                    XLtype = "URI - URL"
                    XLrole = "Embedded_Link"
                if re.match(r"Sandbox report links", line):
                    XLtype = "URI - URL"
                    XLrole = "Embedded_Link"
                if re.match(r"Other hyperlinks", line):
                    XLtype = "URI - URL"
                    XLrole = "Embedded_Link"
                if re.match(r"Downloaded files names and md5s", line):
                    XLtype = ""
                    XLrole = "Attachment"
                if re.match(r"File name", line):
                    XLtype = "File - Name"
                    XLrole = "Attachment"
                if re.match(r"File md5", line):
                    XLtype = "Hash - MD5"
                    XLrole = "Attachment"
                if re.search("</pre>", line): #Switch printing mode OFF
                    P = False
                if P == True: #Print mode is turned ON later, allows JtR to pass over useless lines in the beginning
                    if len(line)>2: #Just in case
                        if re.match(r"http", line): #Seeks hardest items to parse first, links
                            domSeg = line #Domain segment (blah.com)
                            URLSeg = line #URL segment (/blah/d/blah.html)
                            for item in domList:
                                if line.find(item) > 0:
                                    domSeg = re.sub(r"http://",r"", domSeg) #Purifies domain segment
                                    domSeg = re.sub(r"https://",r"", domSeg)
                                    domSeg = re.sub(r"www.",r"", domSeg)
                                    SL = domSeg.index("/") #Slash index, finds where domain ends
                                    domSeg = domSeg[:SL] #Slices domain out
                                    print(r"{}``{}``{}``{}".format(domSeg,"URI - Domain Name","",""), file=fileW)
                                    LP += 1
                            print(r"{}``{}``{}``{}".format(line,XLtype,XLrole,""), file=fileW) #Prints original full line
                            LP += 1
                            URLSeg = re.sub(r"http://",r"", URLSeg) #Purifies URL segment
                            URLSeg = re.sub(r"https://",r"", URLSeg) 
                            URLSeg = re.sub(r"www.",r"", URLSeg)
                            URLSeg = re.sub(domSeg,r"", URLSeg) #Rips out domain
                            iterate = True
                            while iterate == True:
                                last = URLSeg.rfind("/")
                                if last > 0:
                                    URLSeg = URLSeg[:last]
                                    print(r"{}``{}``{}``{}".format(URLSeg,XLtype,XLrole,line), file=fileW)
                                    LP += 1
                                else:
                                    iterate = False
                        if re.search(r"@", line): #search for emails
                            line = re.sub(r'[<">]', r"", line) #remove excess/nonsense characters, inverted the " and ' for capturing "
                            emailName = line.rsplit(None, 1)[0]
                            emailAdd = line.rsplit(None, 1)[-1]
                            print(r"{}``{}``{}``{}".format(emailAdd,"Address - e-mail","Sender_Address",""), file=fileW)
                            LP += 1
                            if emailAdd == emailName:
                                pass
                            else:
                                print(r"{}``{}``{}``{}".format(emailName,"Email Header - String","Sender_Name",emailAdd), file=fileW)
                                LP += 1
                        else:
                            ScanP = False #Scan print boolean
                            for item in extList:
                                if re.search(item, line): #Search for file extensions and seperate file name from MD5s
                                    II = line.index(":") #Item index
                                    print(r"{}``{}``{}``{}".format(line[:II],"File - Name",XLrole,line[II+1:]), file=fileW)
                                    print(r"{}``{}``{}``{}".format(line[II+1:],"Hash - MD5",XLrole,""), file=fileW)
                                    LP += 2
                                    ScanP = True #Line has been printed
                                else:
                                    pass
                            if ScanP == False: #Line was not printed, print line
                                print(r"{}``{}``{}``{}".format(line,XLtype,XLrole,""), file=fileW)
                                LP += 1
                if re.match("<pre>", line): #Switch printing mode on
                    P = True
                    if len(line) > 8:
                        print(r"{}``{}``{}``{}".format(line[5:],XLtype,XLrole,""), file=fileW) #In case of info after <pre>, print line
                        LP += 1
                if re.match("Subject:", line): #Subject line has variable parsing issues, corrected here
                    P = True
                L += 1
            except:
                Cake = True
                pass
    fileW.close()
    print(str(L-1) + " lines parsed, " + str(LP) + " lines printed")

def CK(): #CopyKiller
    global lineList
    print("-----CK START-----")
    ignoreCpy = False
    maxL = sum(1 for line in open(Results))
    print(str(maxL) + " lines to be parsed by CK")
    L = 1 #Current line
    LP = 0 #Lines printed
    CpyK = 0 #Copies killed
    lineList = [] #Line aggregation
    lineAgg = open(LineAgg, "w")
    while L != maxL+1:
        line = linecache.getline(Results, L)
        if re.search(r"�", line) or re.search(r"�", line):
            line = "UNREADABLE CHARACTERS``UNR``UNR``UNR"
        sline = line[:-1].split("``") #Split line     
        if sline[0] in lineList:
            sll = len(sline[0]) #Split line Length
            for item in extList:
                if re.search(item[:-1], sline[0][sll-5:]): #Searches for extension, but only looks in the last 5
                    ignoreCpy = True #                  characters, hoping to not catch URLs with extensions present
            if ignoreCpy == True:
                lineList.append(sline[0])
                print(r"{}".format(line[:-1]), file=lineAgg)
                LP += 1
                ignoreCpy = False
            else:
                CpyK += 1
        else:
            if re.match(r"UNREADABLE", line):
                pass
            else:
                lineList.append(sline[0])
                print(r"{}".format(line[:-1]), file=lineAgg)
                LP += 1
        L += 1
    fileW = open(Results, "w")
    for item in lineList:
        print(r"{}".format(item), file=fileW)
    fileW.close()
    print(str(CpyK) + " copies killed, " + str(LP) + " unique lines printed")

def export():
    print("---EXPORT START---")
    XLSname = ("Files/"+SP.get()+".xls")
    #Formats
    header = xlwt.easyxf("font: name Calibri, bold on")
    inputs = xlwt.easyxf("font: name Calibri")
    #Workbook creation
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Indicators")
    #Information Input
    #Headers
    #       (Row,Col,Data,Format)
    ws.write(0, 0, "Indicator", header)
    ws.write(0, 1, "Type", header)
    ws.write(0, 2, "Comment", header)
    ws.write(0, 3, "Role", header)
    ws.write(0, 4, "Phase", header)
    ws.write(0, 5, "Campaign", header)
    ws.write(0, 6, "Campaign-Description", header)
    ws.write(0, 7, "Campaign-Confidence", header)
    ws.write(0, 8, "Confidence", header)
    ws.write(0, 9, "Impact", header)
    ws.write(0, 10, "Activity-Start", header)
    ws.write(0, 11, "Activity-End", header)
    ws.write(0, 12, "Activity-Description", header)
    ws.write(0, 13, "Bucket", header)
    ws.write(0, 14, "Bucket 1", header)
    ws.write(0, 15, "Relationship-Type", header)
    ws.write(0, 16, "Relationship", header)
    ws.write(0, 17, "Status", header)
    ws.write(0, 18, "RT Ticket", header)
    ws.write(0, 19, "Source", header)
    ws.write(0, 20, "Reference", header)
    #Indicators
    #Static indicator attributes
    Comment = CBE.get("1.0",END)
    Phase = "Delivery"
    Campaign = "zzUnknown"
    Campaign_Description = ""
    Campaign_Confidence = "medium"
    Confidence = "medium"
    Impact = "low"
    Activity_Start = ""
    Activity_End = ""
    Activity_Description = ""
    Bucket = "3000.0-Phishing"
    Bucket_1 = ""
    Relationship_Type = "Related_To"
    Status = "Analyzed"
    RT_Ticket = RT.get()
    Source = "GE IA Intelligence"
    Reference = "https://imweb.corporate.ge.com/wiki/index.php/Category:"+SP.get()
    #Dynamic indicator attributes
    maxL = sum(1 for line in open(LineAgg))
    print(str(maxL) + " lines to be exported")
    L = 1 #Current line
    LP = 0 #Lines printed
    for line in open(LineAgg):
        while L != maxL+1: #Iterate through lines in LineAgg text
            line = linecache.getline(LineAgg, L)
            line = line[:-1].split("``")
            Indicator = line[0]
            Type = line[1]
            Role =  line[2]
            Relationship = line[3]
            #Begin printing attributes to file
            ws.write(L, 0, Indicator, inputs)
            ws.write(L, 1, Type, inputs)
            ws.write(L, 2, Comment, inputs)
            ws.write(L, 3, Role, inputs)
            ws.write(L, 4, Phase, inputs)
            ws.write(L, 5, Campaign, inputs)
            ws.write(L, 6, Campaign_Description, inputs)
            ws.write(L, 7, Campaign_Confidence, inputs)
            ws.write(L, 8, Confidence, inputs)
            ws.write(L, 9, Impact, inputs)
            ws.write(L, 10, Activity_Start, inputs)
            ws.write(L, 11, Activity_End, inputs)
            ws.write(L, 12, Activity_Description, inputs)
            ws.write(L, 13, Bucket, inputs)
            ws.write(L, 14, Bucket_1, inputs)
            ws.write(L, 15, Relationship_Type, inputs)
            ws.write(L, 16, Relationship, inputs)
            ws.write(L, 17, Status, inputs)
            ws.write(L, 18, RT_Ticket, inputs)
            ws.write(L, 19, Source, inputs)
            ws.write(L, 20, Reference, inputs)
            L += 1
    print(str(L-1)+" lines exported to "+XLSname)
    wb.save(XLSname)
    print("Ticket #"+RT.get()+" successfully saved to "+XLSname)
    os.chdir("Files")
    os.startfile(XLSname[6:])
    os.chdir("../")

def cleanup(): #Remove reference files
    os.remove("Files/Results.txt")
    os.remove("Files/LineAgg.txt")

def errormsg1(): #Ticket Error
    popup = Tk()
    popup.wm_title("Error!")
    try:
        popup.wm_iconbitmap(JtR)
    except:
        pass
    popup.geometry("%dx%d+%d+%d" % (GUI_W, GUI_H, ((GUI_X)+(sw*(Screen-1))), GUI_Y))
    label = Label(popup, text="Invalid Ticket Number!", font="Arial 14 bold").pack(side="top", fill="x", pady=(40,0))
    button = Button(popup, width=7, height=3, text="OK", command=popup.destroy).pack(pady=(0,30))

def errormsg2(): #SP ID Error
    popup = Tk()
    popup.wm_title("Error!")
    try:
        popup.wm_iconbitmap(JtR)
    except:
        pass
    popup.geometry("%dx%d+%d+%d" % (GUI_W, GUI_H, ((GUI_X)+(sw*(Screen-1))), GUI_Y))
    label = Label(popup, text="Invalid SP ID Number!", font="Arial 14 bold").pack(side="top", fill="x", pady=(40,0))
    button = Button(popup, width=7, height=3, text="OK", command=popup.destroy).pack(pady=(0,30))

def remindermsg(): #Convert to CSV reminder
    popup = Tk()
    popup.wm_title("Reminder!")
    try:
        popup.wm_iconbitmap(JtR)
    except:
        pass
    popup.geometry("%dx%d+%d+%d" % (GUI_W, GUI_H, ((GUI_X)+(sw*(Screen-1))), GUI_Y))
    label = Label(popup, text="Don't forget to save as a CSV!", font="Arial 14 bold").pack(side="top", fill="x", pady=(40,0))
    button = Button(popup, width=7, height=3, text="OK", command=popup.destroy).pack(pady=(0,30))

def start(): #Run program, called from Start button on GUI
    if RT.get() == "":
        errormsg1()
    elif len(SP.get()) < 15:
        errormsg2()
    else:
        JtR()
        CK()
        export()
        cleanup()
        remindermsg()

#Initiates GUI with SP ID already gathered
INITLINENUM = 0
while INITLINENUM < 10:
    INITLINE = linecache.getline(TBP, INITLINENUM)
    if re.search(r"EventName::", INITLINE):
        SP.set(INITLINE[13:28])
        break
    else:
        INITLINENUM += 1
    
LSPACER1 = Label(root, width=1, bg="black").grid(row=0, column=0)
LSPACER2 = Label(root, width=1, bg="black").grid(row=1, column=0)
RTL = Label(root, text="Ticket Number", bg="black", fg="white", font="Arial 10 bold").grid(row=0, column=1, columnspan=3, sticky=E+W)
RTE = Entry(root, textvariable=RT).grid(row=1, column=1, columnspan=3, sticky=E+W)
MSPACER1 = Label(root, width=1, bg="black").grid(row=0, column=4, padx=3)
MSPACER2 = Label(root, width=1, bg="black").grid(row=1, column=4, padx=3)
SPL = Label(root, text="SP ID Number", bg="black", fg="white", font="Arial 10 bold").grid(row=0, column=5, columnspan=3, sticky=E+W)
SPE = Entry(root, textvariable=SP).grid(row=1, column=5, columnspan=3, sticky=E+W)
RSPACER1 = Label(root, width=1, bg="black").grid(row=0, column=8)
RSPACER2 = Label(root, width=1, bg="black").grid(row=1, column=8)
CBL = Label(root, width=4, text="SP Comment Box", bg="black", fg="white", font="Arial 12 bold")
CBL.grid(row=2, column=0, columnspan=9, sticky=E+W, pady=(5,0), padx=3)
CBE = Text(root, wrap=WORD)
CBE.config(width=45, height=5)
CBE.grid(row=3, column=0, columnspan=9, pady=(0,5), padx=3)
startB = Button(root, text="Start", command=start).grid(row=9, column=2, columnspan=5, sticky=E+W, padx=3)

#Without following line, GUI doesn't work when compiled to .exe format!
#root.mainloop()
