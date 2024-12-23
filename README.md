# uniquify_pictures
A simple utility to identify and remove duplicate pictures in a folder through hashing comparison

## Dependencies
- PIL (may be removed in the future)
- openpyxl
- tqdm

## Command line options:
### -r
If specified, will generate an xlsx spreadsheet report of all the pictures.

### -p
If specified, will copy all files from the input folder to the output folder EXCEPT for duplicate pictures.
Can be used for cleaning up backups from a folder.

### -d
If specified, will generate an xlsx spreadsheet report of only duplicate pictures.

### -i path/to/folder
If specified, uses the path given as the input folder.
If this argument is not given, ./uqinput is used.

### -o path/to/folder
If specified, uses the path given as the output folder.
If this argument is not given, ./uqoutput is used.
