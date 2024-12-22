import os
import sys
import time
import shutil
import hashlib
import operator
from PIL import Image # type: ignore
from openpyxl import Workbook  # type: ignore
from openpyxl.styles import Font, Fill # type: ignore

DEBUG = False
REPORT = False
DUREPORT = False
PLACE = False
INPUTPATH = "./uqinput"
OUTPUTPATH = "./uqoutput"

# scan /uqinput and get a list of pictures
# sort the pictures by metadata/hash
# traverse the list of pictures,
#    adding each to a new list of unique if not equal to last picture

def scan_pictures(path):
    # scan path for all the pictures
    # /uqinput is used for this program

    raw_locations = []
    locations = []
    # traverse folder and get all the files
    print("Scanning "+path+"...")
    for (root, dirs, file) in os.walk(path):
        for f in file:
            raw_locations.append(os.path.join(root, f))
    
    # check for non-pictures
    raw_length = len(raw_locations)
    print(str(raw_length) + " files found.")
    print("Keeping image files and ignoring other files...")
    for i in range(raw_length):
        if check_image(raw_locations[i]):
            locations.append(raw_locations[i])
        print("["+str(i)+"/"+str(raw_length)+"]", end="\r")
    print(str(raw_length)+"/"+str(raw_length)+" checked; "+str(len(locations))+" pictures in "+path)

    return locations

def check_image(f):
    try:
        # open, test, and close
        a = Image.open(f)
        a.tobytes()
        a.close()
        return True
    except:
        return False

def ikey(f):
    # hash images for sorting
    return hashlib.sha256(Image.open(f).tobytes()).hexdigest()

def cache_ikey(data):
    # takes in a list of paths (should be check_imaged) 
    # and outputs them as [ (location, ikey), etc. ]
    return [(loc, ikey(loc)) for loc in data]

def report(sorted_data, tstamp, du=False):
    # takes in a list of paths sorted by image hashes and writes them to xlsx
    print("Generating xlsx " + ("duplicates " if du else "") + "report...")

    wb = Workbook()
    ws = wb.active

    ws.cell(row=1, column=1, value="Filename")
    ws.cell(row=1, column=2, value="Path")
    ws.cell(row=1, column=3, value="Link")

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 60

    ws['A1'].font = Font(bold=True)
    ws['B1'].font = Font(bold=True)
    ws['C1'].font = Font(bold=True)

    if du:
        ws.cell(row=1, column=4, value="Note: Only duplicates in this file.")
        ws.column_dimensions['D'].width = 80
        ws['D1'].font = Font(bold=True)

    cur_row = 2
    len_data = len(sorted_data)

    # handle first element. 
    # if du, only put first element if it is a duplicate
    if (not du) or (sorted_data[0][1] == sorted_data[1][1]):
        real = os.path.realpath(sorted_data[0][0])
        fp = os.path.split(real);
        ws.cell(row=2, column=1, value=fp[1])
        ws.cell(row=2, column=2, value=fp[0])
        ws.cell(row=2, column=3, value='=HYPERLINK("'+real+'")')

        cur_row += 1
    print("[1/"+str(len_data)+"]", end="\r")
    
    if du:
        for i in range(1, len_data):
            if sorted_data[i-1][1] == sorted_data[i][1]:
                # confirmed to be a duplicate
                real = os.path.realpath(sorted_data[i][0])
                fp = os.path.split(real);
                ws.cell(row=cur_row, column=1, value=fp[1])
                ws.cell(row=cur_row, column=2, value=fp[0])
                ws.cell(row=cur_row, column=3, value='=HYPERLINK("'+real+'")')
                cur_row += 1
            else:
                # either start of a new chain or unique
                if i == len_data - 1:
                    # the final element, and not equal to i-1
                    # it is unique
                    pass
                elif sorted_data[i+1][1] == sorted_data[i][1]:
                    # start of a new chain
                    cur_row += 1
                    real = os.path.realpath(sorted_data[i][0])
                    fp = os.path.split(real);
                    ws.cell(row=cur_row, column=1, value=fp[1])
                    ws.cell(row=cur_row, column=2, value=fp[0])
                    ws.cell(row=cur_row, column=3, value='=HYPERLINK("'+real+'")')
                    cur_row += 1
                else:
                    # unique
                    pass
            print("["+str(i+1)+"/"+str(len_data)+"]", end="\r")

    else:
        for i in range(1, len_data):
            if sorted_data[i-1][1] != sorted_data[i][1]:
                cur_row += 1 # insert empty line
            real = os.path.realpath(sorted_data[i][0])
            fp = os.path.split(real);
            ws.cell(row=cur_row, column=1, value=fp[1])
            ws.cell(row=cur_row, column=2, value=fp[0])
            ws.cell(row=cur_row, column=3, value='=HYPERLINK("'+real+'")')
            cur_row += 1
            print("["+str(i+1)+"/"+str(len_data)+"]", end="\r")

    report_name = os.path.join(OUTPUTPATH, "uqreport_"+("duplicates_" if du else "")+str(tstamp)+".xlsx")
    wb.save(report_name)
    print(str(len_data)+"/"+str(len_data)+" entered into report at "+report_name)


def sorted_cleaner(sorted_data):
    # takes in a list of paths sorted by image hashes and then uniquifies them
    print("Uniquifying pictures...")
    unique = [sorted_data[0]]
    len_data = len(sorted_data)
    for i in range(1, len(sorted_data)):
        if sorted_data[i-1][1] != sorted_data[i][1]:
            unique.append(sorted_data[i])
        print("["+str(i)+"/"+str(len_data)+"]", end="\r")
    print(str(len(unique))+"/"+str(len_data)+" pictures are unique")
    return unique

    
def clean(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                # if file, remove
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                # if dir, also remove
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))
        
def main():
    print("\n\u2588 uq.py: Remove duplicate pictures. \u2588\n")

    if PLACE or REPORT:
        print("Cleaning "+OUTPUTPATH+"...")
        clean(OUTPUTPATH)

    sloc = sorted(cache_ikey(scan_pictures(INPUTPATH)),\
        key=operator.itemgetter(1))

    timestamp = int(1000 * time.time())

    if REPORT:
        report(sloc, timestamp)
    
    if DUREPORT:
        report(sloc, timestamp, du=True)

    if PLACE:
        unique = sorted_cleaner(sloc)
        print("Placing unique pictures into "+OUTPUTPATH)
        for file in unique:
            rel_path = os.path.relpath(file[0], INPUTPATH)
            dest = os.path.join(OUTPUTPATH, rel_path)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(file[0], dest)

def debug():
    pass

if __name__ == "__main__":
    if "-r" in sys.argv:
        REPORT = True

    if "-p" in sys.argv:
        PLACE = True
    
    if "-d" in sys.argv:
        DUREPORT = True
    
    if "-i" in sys.argv:
        INPUTPATH = sys.argv[sys.argv.index("-i")+1]
    else:
        print("No input directory specified with -i; defaulting to ./uqinput")
    
    if "-o" in sys.argv:
        OUTPUTPATH = sys.argv[sys.argv.index("-o")+1]
    else:
        print("No output directory specified with -o; defaulting to ./uqoutput")
    
    if (not REPORT) and (not PLACE):
        print("Warning: no recognized arguments specified.")
        print("Use -r to generate a report of all files.")
        print("Use -d to generate a report of duplicates.")
        print("Use -p to copy unique pictures into output.")
        print("Use -i [input path] to specify an input path.")
        print("Use -o [output path] to specify an output path.")

    if DEBUG:
        print("Running in debug mode.")
        debug()
    else:
        main()
