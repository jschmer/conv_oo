import os
import fnmatch
import sys
import argparse

parser = argparse.ArgumentParser(description="""
Convert media files (for example odp -> pdf)
>>> The soffice executable of LibreOffice (tested) or OpenOffice (untested) has to be in PATH!
""",
formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument("-s", "--source", action='store', nargs=1, default=".", type=str, help="Source directory where your media files are stored")
parser.add_argument("-ff", "--file-filter", action='store', nargs=1, default="*.odp", type=str, help="Convert only the files in source that matches this filter, for example '*.odp'")
parser.add_argument("-o", "--out-dir", action='store', nargs=1, default=".", type=str, help="Directory to hold the converted files\nIs relative to source if a relative path is given")
parser.add_argument("-t", "--type", action='store', nargs=1, default="pdf", type=str, help="Media type to convert to, for example 'pdf'")

args = parser.parse_args()
    
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


def quote(path):
  if path[0] != path[len(path)-1]:
    return '"' + path + '"'
  else:
    return path

if which("soffice.exe") == None:
    parser.print_help()
    sys.exit(2)

# conversion command: "soffice --headless --convert-to pdf --outdir #outdir# #file1# #file2# ... #fileN#"

source = os.path.abspath(args.source[0])
if not os.path.isdir(source):
    print("ERROR: source is not a directory: args.source")
    sys.exit(-1)

# get all odp files in base_dir
files = []
for root, dirnames, filenames in os.walk(source):
  for filename in fnmatch.filter(filenames, args.file_filter[0]):
      files.append('"' + os.path.join(root, filename) + '"')

if len(files) < 1:
    print("Nothing to convert, aborting...")
    sys.exit(0)

# prepare outdir relative to base_dir
if os.path.isabs(args.out_dir[0]):
  outdir = quote(args.out_dir[0])
else:
  outdir = quote(os.path.abspath(os.path.join(source, args.out_dir[0])))

cmd = "soffice.exe --headless --convert-to " + args.type[0] + " --outdir " + outdir + " " + " ".join(files)
print(cmd)

os.system(cmd)
