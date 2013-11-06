import os
import fnmatch
import sys
import getopt

def usage():
    print('LibreOffice soffice.exe must be in path!')
    print('')
    print('convert_oo.py --dir "DV-Recht" --in-file-filter "*.odp" --out-type "pdf" --out-dir "pdf"')
    print('')
    print('--out-dir is relative to --dir if it\'s a relative path')

def quote(path):
  if path[0] != path[len(path)-1]:
    return '"' + path + '"'
  else:
    return path
    
def which(program):
    '''
    Taken from http://stackoverflow.com/a/377028
    '''
    def is_exe(fpath):
        return os.path.isfile(fpath) and os.access(fpath, os.X_OK)

    fpath, fname = os.path.split(program)
    if fpath:
        if is_exe(program):
            return program
    else:
        for path in os.environ["PATH"].split(os.pathsep):
            path = path.strip('"')
            exe_file = os.path.join(path, program)
            if is_exe(exe_file):
                return exe_file

    return None

if which("soffice.exe") == None:
    usage()                         
    sys.exit(2)    

# default arguments
base_dir = "."
input_file_filter = "*.odp"
output_format = "pdf"
outdir = "."
     
try:                                
    opts, args = getopt.getopt(sys.argv[1:], "hd:i:t:o:", ["help", "dir=", "in-file-filter=", "out-type=", "out-dir="])
except getopt.GetoptError:          
    usage()                         
    sys.exit(2)                     
for opt, arg in opts:
    if opt in ("-h", "--help"):
        usage()
        sys.exit(1)
    elif opt in ("-d", "--dir"):
        base_dir = os.path.abspath(arg)
    elif opt in ("-i", "--in-file-filter"):
        input_file_filter = arg              
    elif opt in ("-t", "--out-type"):
        output_format = arg          
    elif opt in ("-o", "--out-dir"):
        outdir = arg

# conversion command: "soffice --headless --convert-to pdf --outdir #outdir# #file1# #file2# ... #fileN#"

# get all odp files in base_dir
files = []
for root, dirnames, filenames in os.walk(base_dir):
  for filename in fnmatch.filter(filenames, input_file_filter):
      files.append('"' + os.path.join(root, filename) + '"')

# prepare outdir relative to base_dir
if os.path.isabs(outdir):
  outdir = '"' + outdir + '"'
else:
  outdir = '"' + os.path.abspath(os.path.join(base_dir, outdir)) + '"'

cmd = "soffice.exe --headless --convert-to " + output_format + " --outdir " + outdir + " " + " ".join(files)

os.system(cmd)
