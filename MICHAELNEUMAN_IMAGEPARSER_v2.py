import subprocess,time,os,win32file,win32con,datetime,dateutil.parser,shutil,csv,sys
from tqdm import tqdm
maxInt = sys.maxsize
print("CALCULATING MAXSIZE...CURRENTLY "+str(maxInt),end="...")
while True:
    # decrease the maxInt value by factor 10
    # as long as the OverflowError occurs.
    try:
        csv.field_size_limit(maxInt)
        break
    except OverflowError:
        maxInt = int(maxInt/10)
print("FINISHED WITH MAXSIZE " + str(maxInt))
def get_jumplist_filepaths():
    jumplist_filepaths = set()
    for filepath in inode_data.keys():
        if any([jumplist_extension in filepath for jumplist_extension in set(('automaticdestinations-ms','customdestinations-ms'))]):
            jumplist_filepaths.add(filepath)
    return jumplist_filepaths

def get_lnk_filepaths():
    lnk_filepaths = set()
    for filepath in inode_data.keys():
        if filepath[-4:] == '.lnk':
            lnk_filepaths.add(filepath)
    return lnk_filepaths

def get_evtx_filepaths():
    evtx_filepaths = set()
    for filepath in inode_data.keys():
        if filepath[-5:] == '.evtx':
            evtx_filepaths.add(filepath)
    return evtx_filepaths

def get_pf_filepaths():
    pf_filepaths = set()
    for filepath in inode_data.keys():
        if filepath[-3:] == '.pf':
            pf_filepaths.add(filepath)
    return pf_filepaths

def get_recyclebin_files():
    rb_filepaths = set()
    for filepath in inode_data.keys():
        if filepath[:12] == '$recycle.bin' and '$i' in filepath.split('/')[-1]:
            rb_filepaths.add(filepath)
    return rb_filepaths

def get_registry_files():
    return_set = set((
        'windows/system32/config/sam','windows/system32/config/sam.log1','windows/system32/config/sam.log2',
        'windows/system32/config/security','windows/system32/config/security.log1','windows/system32/config/security.log2',
'windows/system32/config/software.log1','windows/system32/config/software.log2','windows/system32/config/software','windows/system32/config/system',
'windows/system32/config/system.log1','windows/system32/config/system.log2'
    ))
    for filepath in inode_data.keys():
        if any([extension in filepath for extension in set(('ntuser.dat','ntuser.dat.log1','ntuser.dat.log2','usrclass.dat','usrclass.dat.log1','usrclass.dat.log2'))]):
            return_set.add(filepath)
    return return_set

def get_shellbag_files():
    return_set = set()
    for filepath in inode_data.keys():
        if any([extension in filepath for extension in set(('ntuser.dat','ntuser.dat.log1','ntuser.dat.log2','usrclass.dat','usrclass.dat.log1','usrclass.dat.log2'))]):
            return_set.add(filepath)
    return return_set

def get_relevant_sql_files():
    return_set,filenames = set(),set()
    for file in os.listdir(OTHERTOOLS_DIR + "Maps"):
        if '.smap' in file.lower():
            filenames.add(
                open(OTHERTOOLS_DIR + "Maps\\" + file, 'r', encoding='latin-1').read().split(
                    'FileName: ')[1].split('\n')[0].split('#')[0].strip())
    for filepath in inode_data.keys():
        if any([filename.lower() in filepath for filename in filenames]):
            return_set.add(filepath)
    return return_set

def get_win10timeline_files():
    return_set = set()
    for filepath in inode_data.keys():
        if filepath[-18:]=='activitiescache.db':
            return_set.add(filepath)
    return return_set

def get_ual_files():
    return_set = set()
    for filepath in inode_data.keys():
        if 'Windows/System32/LogFiles/SUM'.lower() in filepath:
            return_set.add(filepath)
    return return_set

def get_srum_files():
    return_set = set()
    for filepath in inode_data.keys():
        if filepath=='windows/system32/sru/srudb.dat' or 'Windows/System32/Config/SOFTWARE'.lower() in filepath:
            return_set.add(filepath)
    return return_set

# define globals
start_time = time.time()
OTHERTOOLS_DIR = "E:\\MICHAELNEUMAN_IMAGEPARSER\\tools\\"
SLEUTHKIT_DIR = OTHERTOOLS_DIR+"sleuthkit-4.12.0\\bin\\"
IMAGE_PATH = "E:\\MICHAELNEUMAN_IMAGEPARSER\\test_output_directory\\Image\\CDRIVE_202303141137.E01"
OUTPUT_DIR = "E:\\MICHAELNEUMAN_IMAGEPARSER\\test_output_directory\\Export\\"
REPORT_DIR = "E:\\MICHAELNEUMAN_IMAGEPARSER\\test_output_directory\\Timeline\\"
RECYCLEBIN_DIR = OUTPUT_DIR

inode_data_file = open("E:\\MICHAELNEUMAN_IMAGEPARSER\\test_output_directory\\inode_temp.txt",'r',encoding='latin-1')#{}
inode_data = {}
amcache_filepaths = set(("Windows/appcompat/Programs/Amcache.hve","Windows/appcompat/Programs/Amcache.hve.LOG1","Windows/appcompat/Programs/Amcache.hve.LOG2"))
shimcache_filepaths = set(("Windows/System32/Config/SYSTEM","Windows/System32/Config/SYSTEM.LOG1","Windows/System32/Config/SYSTEM.LOG2"))
mft_filepaths = set(("$MFT",))
REPORT_SUBDIRS = set(("AMCACHE","SHIMCACHE","MFT","JUMPLISTS","LNK","EVTX","PREFETCH","RECYCLEBIN","REGISTRY","SHELLBAGS","SQL","WIN10TIMELINE","UAL","SRUM"))

# TEMPORARILY USING TEST DATA
#
#print(SLEUTHKIT_DIR+"fls -r -p "+IMAGE_PATH)
#for line in tqdm(subprocess.check_output([SLEUTHKIT_DIR+"fls","-r","-p",IMAGE_PATH]).decode("latin-1").splitlines()):
#    inode_location, file_name = " ".join(line.split(" ")[1:]).split('\t')
#    inode_data[file_name.lower()] = inode_location.replace(':','')
#    inode_data_file.write(file_name.lower()+'\t'+inode_location.replace(':','').replace('*','').strip()+'\n')
#    inode_data_file.flush()
for line in inode_data_file.read().splitlines():
    temp = line.split('\t')
    inode_data[temp[0]] = temp[1].replace('*','').strip()

jumplist_filepaths = set((get_jumplist_filepaths()))
lnk_filepaths = set((get_lnk_filepaths()))
evtx_filepaths = set((get_evtx_filepaths()))
pf_filepaths = set((get_pf_filepaths()))
recyclebin_filepaths = set((get_recyclebin_files()))
registry_filepaths = set((get_registry_files()))
shellbag_filepaths = set((get_shellbag_files()))
sql_filepaths = set((get_relevant_sql_files()))
win10_filepaths = set((get_win10timeline_files()))
ual_filepaths = set((get_ual_files()))
srum_filepaths = set((get_srum_files()))
#files_to_acquire = set(x.lower() for x in (amcache_filepaths | shimcache_filepaths | mft_filepaths | jumplist_filepaths | lnk_filepaths | evtx_filepaths | pf_filepaths | recyclebin_filepaths | registry_filepaths | shellbag_filepaths | sql_filepaths | win10_filepaths | ual_filepaths | srum_filepaths))
files_to_acquire = set() #sql_filepaths
local_filepaths = {}

def adjust_timestamps(filepath,inode_location):
    try:
        temp = subprocess.check_output([SLEUTHKIT_DIR + "istat", IMAGE_PATH, inode_location]).decode("latin-1").split("$STANDARD_INFORMATION Attribute Values:")[1].split("$FILE_NAME Attribute Values:")[0].strip()
    except:
        pass
    if len(temp)>0:
        new_cma_timestamps = [None, None, None]
        for line in temp.splitlines():
            if 'Created:' in line:
                new_cma_timestamps[0] = line.split(".")[0].replace("Created:", "").strip()
            elif 'File Modified:' in line:
                new_cma_timestamps[1] = line.split(".")[0].replace("File Modified:", "").strip()
            elif 'Accessed:' in line:
                new_cma_timestamps[2] = line.split(".")[0].replace("Accessed:", "").strip()
        if None not in new_cma_timestamps:
            now = time.localtime()
            # handle datetime.datetime parameters
            ctime = time.mktime(dateutil.parser.parse(new_cma_timestamps[0]).timetuple())
            mtime = time.mktime(dateutil.parser.parse(new_cma_timestamps[1]).timetuple())
            atime = time.mktime(dateutil.parser.parse(new_cma_timestamps[2]).timetuple())
            # adjust for day light savings
            ctime += 3600 * (now.tm_isdst - time.localtime(ctime).tm_isdst)
            mtime += 3600 * (now.tm_isdst - time.localtime(mtime).tm_isdst)
            atime += 3600 * (now.tm_isdst - time.localtime(atime).tm_isdst)
            # change time stamps
            winfile = win32file.CreateFile(
                filepath, win32con.GENERIC_WRITE,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None, win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL, None)
            win32file.SetFileTime(winfile,
                                  datetime.datetime.utcfromtimestamp(ctime).replace(tzinfo=datetime.timezone.utc),
                                  datetime.datetime.utcfromtimestamp(mtime).replace(tzinfo=datetime.timezone.utc),
                                  datetime.datetime.utcfromtimestamp(atime).replace(tzinfo=datetime.timezone.utc))
            winfile.close()

for file in tqdm(files_to_acquire):
    file_name,output_directory_local = file,OUTPUT_DIR
    if file_name[:12]=='$recycle.bin':
        output_directory_local = RECYCLEBIN_DIR
    for i in range(1,len(file.split('/')[:-1])+1):
        try:
            os.mkdir(output_directory_local + '/'.join(file.split('/')[:i]))
        except:
            pass
    try:
        with open(output_directory_local + file_name, 'wb') as outfile:
            outfile.write(subprocess.check_output([SLEUTHKIT_DIR+"icat", IMAGE_PATH, inode_data[file]]))
        if file_name[:12]!='$recycle.bin':
            adjust_timestamps(output_directory_local + file_name, inode_data[file])
        local_filepaths[file] = output_directory_local + file_name
    except:
        pass

def amcache_report():
    temp_subprocess_string = [OTHERTOOLS_DIR+"AmcacheParser.exe", "-f",local_filepaths["Windows/appcompat/Programs/Amcache.hve".lower()],"--csv",REPORT_DIR+"AMCACHE"]
    subprocess.run(temp_subprocess_string)
    if len(os.listdir(REPORT_DIR+"AMCACHE"))==0:
        subprocess.run(temp_subprocess_string + ["--nl"])

def shimcache_report():
    temp_subprocess_string = [OTHERTOOLS_DIR + "AppCompatCacheParser.exe", "-f",local_filepaths["Windows/System32/Config/SYSTEM".lower()], "--csv",REPORT_DIR + "SHIMCACHE"]
    subprocess.run(temp_subprocess_string)
    if len(os.listdir(REPORT_DIR + "SHIMCACHE")) == 0:
        subprocess.run(temp_subprocess_string + ["--nl"])

def mft_report():
    subprocess.run([OTHERTOOLS_DIR+ "MFTECmd.exe", "-f", local_filepaths["$MFT".lower()],"--csv", REPORT_DIR+"MFT"])

def jumplist_report():
    subprocess.run([OTHERTOOLS_DIR+ "JLECmd.exe", "-d", OUTPUT_DIR,"--csv", REPORT_DIR+ "JUMPLISTS"])

def lnk_report():
    subprocess.run([OTHERTOOLS_DIR + "LECmd.exe", "-d", OUTPUT_DIR, "--csv", REPORT_DIR + "LNK"])

def evtx_report():
    subprocess.run([OTHERTOOLS_DIR + "EvtxECmd.exe", "-d", OUTPUT_DIR, "--csv", REPORT_DIR + "EVTX"])

def prefetch_report():
    subprocess.run([OTHERTOOLS_DIR + "PECmd.exe", "-d", OUTPUT_DIR, "--csv", REPORT_DIR + "PREFETCH"])

def recyclebin_report():
    f_name = 1
    for path,subdirs,files in os.walk(RECYCLEBIN_DIR+ '$recycle.bin'):
        for name in files:
            subprocess.run([OTHERTOOLS_DIR + "RBCmd.exe", "-f", os.path.join(path,name), "--csv", REPORT_DIR + "RECYCLEBIN","--csvf",str(f_name)+'.csv'])
            f_name+= 1
    final_report = open(REPORT_DIR + "RECYCLEBIN\\COMBINED_RECYCLEBIN_REPORT.tsv",'w',encoding='latin-1')
    final_report.write("\t".join(["SourceName","FileType","FileName","FileSize","DeletedOn"])+'\n')
    for recyclebin_filereport in os.listdir(REPORT_DIR + "RECYCLEBIN"):
        if recyclebin_filereport!='COMBINED_RECYCLEBIN_REPORT.tsv':
            try:
                temp = open(REPORT_DIR + "RECYCLEBIN\\"+recyclebin_filereport,'r',encoding='latin-1').read().splitlines()[1].split(',')
                temp[0] = temp[0].replace(OUTPUT_DIR,'')
                final_report.write('\t'.join(temp)+'\n')
            except:
                pass
            final_report.flush()
            os.remove(REPORT_DIR + "RECYCLEBIN\\"+recyclebin_filereport)
def create_report_directories():
    for dirname in REPORT_SUBDIRS:
        try:
            os.mkdir(REPORT_DIR + dirname)
        except:
            pass
def cleanup():
    for path,subdirs,files in os.walk(OUTPUT_DIR):
        for name in files:
            try:
                os.remove(os.path.join(path,name))
            except:
                pass
        for subdir in subdirs:
            try:
                shutil.rmtree(os.path.join(path,subdir))
            except:
                pass
    try:
        os.remove(OUTPUT_DIR + "$MFT")
    except:
        pass

def registry_report():
    subprocess.run([OTHERTOOLS_DIR + "RECmd.exe", "-d", OUTPUT_DIR, "--csv", REPORT_DIR + "REGISTRY",'--bn',OTHERTOOLS_DIR + 'BatchExamples\\Kroll_Batch.reb'])

def shellbags_report():
    subprocess.run([OTHERTOOLS_DIR + "SBECmd.exe","-d",OUTPUT_DIR,"--csv",REPORT_DIR+"SHELLBAGS"])

def sql_report():
    subprocess.run([OTHERTOOLS_DIR + "SQLECmd.exe","-d",OUTPUT_DIR,"--csv",REPORT_DIR+"SQL"])
    subprocess.run([OTHERTOOLS_DIR + "BrowsingHistoryView.exe","/HistorySource","2",OUTPUT_DIR,"/stabular",REPORT_DIR+"SQL\\BROWSERREPORT.txt"])

def windows10timeline_report():
    for path, subdirs, files in os.walk(OUTPUT_DIR):
        for name in files:
            if name=='activitiescache.db':
                subprocess.run([OTHERTOOLS_DIR + "WxTCmd.exe","-f",os.path.join(path,name),"--csv",REPORT_DIR+"WIN10TIMELINE"])

def ual_report():
    try:
        subprocess.run([OTHERTOOLS_DIR + "SumECmd.exe","-d",OUTPUT_DIR+'Windows/System32/LogFiles/SUM'.lower(),"--csv",REPORT_DIR+"UAL"])
    except:
        pass
def srum_report():
    subprocess.run([OTHERTOOLS_DIR + "SrumECmd","-f",OUTPUT_DIR+'windows/system32/sru/srudb.dat',"-r",OUTPUT_DIR+'windows/system32/config/SOFTWARE','--csv',REPORT_DIR+"SRUM"])

def generate_fulltimeline():
    fulltimeline = open(REPORT_DIR + 'FULLTIMELINE.tsv','w',encoding='latin-1')
    fulltimeline.write('\t'.join(['TIMESTAMP (UTC)','HOSTNAME','SHORT DESCRIPTION','DESCRIPTION','INDICATOR(HASH/IP)','ARTIFACT'])+'\n')
    for dir in REPORT_SUBDIRS:
        for file in os.listdir(REPORT_DIR + dir):
            temp = generate_timeline_fromsinglefile(dir,REPORT_DIR + dir + '/' + file)
            if len(temp)>0:
                fulltimeline.write(temp)
                fulltimeline.flush()

def generate_timeline_fromsinglefile(report_directory,full_filepath):
    if report_directory=='RECYCLEBIN':
        return generate_recyclebin_timeline(full_filepath)
    elif report_directory=='LNK':
        return generate_lnk_timeline(full_filepath)
    elif report_directory=='AMCACHE':
        return generate_amcache_timeline(full_filepath)
    elif report_directory=='PREFETCH':
        return generate_prefetch_timeline(full_filepath)
    elif report_directory=='SHELLBAGS':
        return generate_shellbags_timeline(full_filepath)
    elif report_directory=='SQL':
        return generate_browser_timeline(full_filepath)
    else:
        return ""

def generate_recyclebin_timeline(full_filepath):
    return_string = ""
    for line in open(full_filepath,'r',encoding='latin-1').read().splitlines()[1:]:
        timeline_entry = [None, None, None, None, None, None]
        line = line.split('\t')
        timeline_entry[0] = line[-1]
        timeline_entry[1] = "Hostname"
        timeline_entry[2] = "File Deletion"
        timeline_entry[3] = "File '"+line[2]+"' recycled"
        timeline_entry[4] = line[0].replace('$recycle.bin\\','').split('\\')[0]
        timeline_entry[5] = '$RECYCLE.BIN'
        return_string += '\t'.join(timeline_entry)+'\n'
    return return_string

def generate_lnk_timeline(full_filepath):
    return_string = ""
    reader, first = csv.reader(open(full_filepath, 'r', encoding='latin-1')), True
    for line in reader:
        if not first:
            try:
                line[0] = line[0].replace('\\\\','\\')
                if '\\users\\' in line[0] and len(line[4])>1 and '\\appdata\\' in line[0]:
                    user = line[0].replace(OUTPUT_DIR,'').split('\\')[1]
                    timeline_entry_created,timeline_entry_modified = [None, None, None, None, None, None],[None, None, None, None, None, None]
                    timeline_entry_created[0],timeline_entry_modified[0] = line[1],line[2]
                    timeline_entry_created[1],timeline_entry_modified[1] = "Hostname","Hostname"
                    timeline_entry_created[2],timeline_entry_modified[2] = "File/Directory Access","File/Directory Access"
                    args_string = ''
                    if len(line[18])>0:
                        args_string = ' '+line[18]
                    timeline_entry_created[3],timeline_entry_modified[3] = "File '" + line[15] + line[17] + args_string +"' first accessed by user "+user+" (size: "+line[7]+")","File '" + line[15] + line[17] + args_string +"' last accessed by user "+user+" (size: "+line[7]+")"
                    timeline_entry_created[4],timeline_entry_modified[4] = user,user
                    timeline_entry_created[5],timeline_entry_modified[5] = line[0].replace(OUTPUT_DIR,''),line[0].replace(OUTPUT_DIR,'')
                    return_string += '\t'.join(timeline_entry_created) + '\n' + '\t'.join(timeline_entry_modified) + '\n'
            except Exception as e:
                print(str(e))
        first = False
    return return_string

def generate_amcache_timeline(full_filepath):
    return_string = ""
    if 'UnassociatedFileEntries' in full_filepath:
        reader, first = csv.reader(open(full_filepath, 'r', encoding='latin-1')), True
        for line in reader:
            if not first and len(line[7])>0:
                timeline_entry = [None, None, None, None, None, None]
                timeline_entry[0] = line[2]
                timeline_entry[1] = "Hostname"
                timeline_entry[2] = "Program Execution"
                timeline_entry[3] = "Program '" + line[5] + "' last executed (size: "+line[10]+";productname: "+line[9]+")"
                hash_string = ""
                if len(line[3])==40:
                    hash_string = line[3]
                timeline_entry[4] = hash_string
                timeline_entry[5] = 'Amcache.hve'
                return_string += '\t'.join(timeline_entry) + '\n'
            first = False
    return return_string

def generate_prefetch_timeline(full_filepath):
    return_string = ""
    if 'PECmd_Output_Timeline' not in full_filepath:
        reader, first = csv.reader(open(full_filepath, 'r', encoding='latin-1')), True
        for line in reader:
            if not first:
                timeline_entry = [None, None, None, None, None, None]
                runtimes = set(line[10:17])
                for runtime in runtimes:
                    if len(runtime)>0:
                        timeline_entry[0] = runtime
                        timeline_entry[1] = "Hostname"
                        timeline_entry[2] = "Program Execution"
                        timeline_entry[3] = "Program '" + line[5] + "' executed (size: " + line[7] + ")"
                        timeline_entry[4] = ""
                        timeline_entry[5] = line[1]
                        return_string += '\t'.join(timeline_entry) + '\n'
            first = False
    return return_string

def generate_shellbags_timeline(full_filepath):
    return_string = ""
    if '!SBECmd_Messages' not in full_filepath:
        reader, first = csv.reader(open(full_filepath, 'r', encoding='latin-1')), True
        for line in reader:
            if not first:
                timeline_entry = [None, None, None, None, None, None]
                first_interacted,last_interacted = line[15],line[16]
                hive_path = open('/'.join(full_filepath.split('/')[:-1])+'/!SBECmd_Messages.txt','r',encoding="latin-1").read().split(full_filepath.replace('/','\\'))[0].split("Finished processing ")[-1].split('\n')[0].strip()
                user = hive_path.split('\\users\\')[1].split('\\')[0]
                if len(first_interacted)>0:
                    timeline_entry[0] = first_interacted
                    timeline_entry[1] = "Hostname"
                    timeline_entry[2] = "File/Directory Access"
                    timeline_entry[3] = "Directory '" + line[4] + "' first interacted by user " + user
                    timeline_entry[4] = user
                    timeline_entry[5] = hive_path.replace(OUTPUT_DIR,'')
                    return_string += '\t'.join(timeline_entry) + '\n'
                if len(last_interacted)>0:
                    timeline_entry[0] = last_interacted
                    timeline_entry[1] = "Hostname"
                    timeline_entry[2] = "File/Directory Access"
                    timeline_entry[3] = "Directory '" + line[4] + "' last interacted by user " + user
                    timeline_entry[4] = user
                    timeline_entry[5] = hive_path.replace(OUTPUT_DIR,'')
                    return_string += '\t'.join(timeline_entry) + '\n'
            first = False
    return return_string

def generate_browser_timeline(full_filepath):
    return_string = ""
    if 'BROWSERREPORT' in full_filepath:
        reader   = csv.reader(open(full_filepath, 'r',encoding="latin-1"))
        for line in reader:
            print(line)
            if not first:
                try:
                    timeline_entry = [None, None, None, None, None, None]
                    timeline_entry[0] = line[2]
                    timeline_entry[1] = "Hostname"
                    timeline_entry[2] = "Browser History"
                    url_string = line[0][:150]
                    if len(line[0])>150:
                        url_string += '...'
                    timeline_entry[3] = "URL '" + url_string+ "' ("+line[1]+") accessed interacted by user " + line[8]
                    timeline_entry[4] = line[8]
                    timeline_entry[5] = line[12].lower().replace('\\\\','\\')
                    return_string += '\t'.join(timeline_entry) + '\n'
                except:
                    pass
            first = False
    return return_string


#create_report_directories()
#amcache_report()
#shimcache_report()
#mft_report()
#jumplist_report()
#lnk_report()
#evtx_report()
#prefetch_report()
#recyclebin_report()
#registry_report()
#shellbags_report()
sql_report()
#windows10timeline_report()
#ual_report()
#srum_report()
generate_fulltimeline()
#generate_last90daystimeline()
#cleanup()
print("PROGRAM TOOK {runtime} SECONDS TO RUN".format(runtime=time.time()-start_time))