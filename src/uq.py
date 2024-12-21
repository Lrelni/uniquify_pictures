import os
import sys
import time
import shutil
import hashlib
from PIL import Image
from openpyxl import Workbook 

DEBUG = False
REPORT = True
PLACE = True
INPUTPATH = "./uqinput2"
OUTPUTPATH = "./uqoutput2"

# scan and get a list of pictures
# sort the pictures by metadata/hash
# traverse the list of pictures,
#    adding each to a new list of unique if not equal to last picture
# generate a report if needed

def scan_pictures(path):
    # scan path for all the pictures
    # /uqinput is used for this program

    locations = []
    # traverse folder and get all the files
    for (root, dirs, file) in os.walk(path):
        for f in file:
            locations.append(os.path.join(root, f))
    return locations

def ikey(f):
    # hash images for sorting
    return hashlib.sha256(Image.open(f).tobytes()).hexdigest()

def uniquify(data):
    # remove duplicates by sorting by hash
    # and keeping if hash if i-1 neq heah of i

    print("Sorting pictures")
    sloc = sorted(data, key=ikey)
    write_to_xlsx(sloc)
    return sorted_cleaner(sloc)

def write_to_xlsx(sorted_data):
    # takes in a list of paths sorted by image hashes and writes them to xlsx
    print("Generating xlsx report")
    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="Filename")
    ws.cell(row=1, column=2, value="Path")
    
    cur_row = 2
    for i in range(1, len(sorted_data)):
        if ikey(sorted_data[i-1]) != ikey(sorted_data[i]):
            cur_row += 1 # insert empty line
        fp = os.path.split(os.path.realpath(sorted_data[i]));
        ws.cell(row=cur_row, column=1, value=fp[1])
        ws.cell(row=cur_row, column=2, value=fp[0])
        cur_row += 1
    wb.save("uqreport_"+str(int(1000 * time.time()))+".xlsx")


def sorted_cleaner(sorted_data):
    # takes in a list of paths sorted by image hashes and then uniquifies them
    print("Uniquifying")
    unique = [sorted_data[0]]
    for i in range(1, len(sorted_data)):
        if ikey(sorted_data[i-1]) != ikey(sorted_data[i]):
            unique.append(sorted_data[i])
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
    print("\nuq.py: Remove duplicate pictures.\n")
    if PLACE:
        print("Cleaning "+OUTPUTPATH)
        clean(OUTPUTPATH)
    print("Reading pictures from "+INPUTPATH)
    sloc = sorted(scan_pictures(INPUTPATH), key=ikey)
    if REPORT:
        write_to_xlsx(sloc)
    if PLACE:
        unique = sorted_cleaner(sloc)
        print("Placing unique pictures into "+OUTPUTPATH)
        for file in unique:
            rel_path = os.path.relpath(file, INPUTPATH)
            dest = os.path.join(OUTPUTPATH, rel_path)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(file, dest)

def debug():
    pass

if __name__ == "__main__":
    if DEBUG:
        print("Running in debug mode.")
        debug()
    else:
        main()
