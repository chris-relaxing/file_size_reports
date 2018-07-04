#-------------------------------------------------------------------------------
# Name:         File Size Report
# Purpose:      Input 2 directory paths to be compared. Also choose the report name.
#
# Author:       Chris Nielsen
# Created:      July 31, 2014
#-------------------------------------------------------------------------------

import os
import math
import time
import datetime
import xlwt as xlwt
from Tkinter import *
import tkMessageBox
import tkFileDialog

start_time = time.time()

report_name = ''
cleaned_old_path = ''
cleaned_new_path = ''
paths_to_compare = []

#-------------------------------------------------------------------------------
class pathDialog(Frame):

    def __init__(self, parent, **kwargs):
        Frame.__init__(self, parent, **kwargs)
        self.parent = parent
        self.parent.title("File Size Report")
        self.pack(fill=BOTH, expand=1)
        self.centerWindow()
        self.initUI()

        parent.protocol("WM_DELETE_WINDOW", self.handler)

    def handler(self):
        if tkMessageBox.askokcancel("Quit?", "Are you sure you want to quit?"):
            self.parent.destroy()
            print "Destoy root window."
            print "Quit main loop."
            sys.exit()

    def initUI(self):

        def loadDir(entry):
            chosen_path = tkFileDialog.askdirectory()
            entry.set(chosen_path)

        def printSelectedDirs(old_path, new_path, report_name_entry):

            global cleaned_old_path
            global cleaned_new_path
            global paths_to_compare
            global report_name
            global Selected_old_path
            global Selected_new_path

            report_name = str(report_name_entry.get())
            print "The name entered for this report is:", report_name

            Selected_old_path = ""
            Selected_old_path = str(old_path.get())
            cleaned_old_path = remove_trailing_slash(Selected_old_path)

            Selected_new_path = ""
            Selected_new_path = str(new_path.get())
            cleaned_new_path = remove_trailing_slash(Selected_new_path)

            if not Selected_old_path:
                ThrowError("Error: No path has been selected.", "", "Error: No path has been selected.")

            if not Selected_new_path:
                ThrowError("Error: No path has been selected.", "", "Error: No path has been selected.")

            # old folder ----
            old_legitfolder = os.path.isdir(Selected_old_path)

            if old_legitfolder:
                print "The Selected_old_path is:", Selected_old_path
                print "Cleaned path is", cleaned_old_path
            else:
                ThrowError("Error: Not a valid path.", "Error: Not a valid path:", cleaned_old_path)

            # new folder ----
            new_legitfolder = os.path.isdir(Selected_new_path)

            if new_legitfolder:
                print "The Selected_new_path is:", Selected_new_path
                print "Cleaned path is", cleaned_new_path
            else:
                ThrowError("Error: Not a valid path.", "Error: Not a valid path:", cleaned_new_path)

            paths_to_compare = [cleaned_old_path, cleaned_new_path]

        def multCommands(old_path, new_path, report_name_entry):
            printSelectedDirs(old_path, new_path, report_name_entry)
            if Selected_old_path and Selected_new_path:
                self.parent.destroy() # This is what destroys the main window

        self.title = 'File Size Comparisons'
        Label(self, text = 'Path to OLD directory to be compared:').grid(row = 0, column = 0, pady=0)

        old_path = StringVar()
        Entry(self, width=100,textvariable = old_path).grid(row = 1, column =0, columnspan = 1, pady=0)
        Button(self, text = "Browse", command = lambda: loadDir(old_path), width = 10).grid(row = 1, column = 1, pady=0)

        new_path = StringVar()
        Label(self, text = 'Path to NEW directory to be compared:').grid(row = 2, column = 0, pady=0)
        Entry(self, width=100,textvariable = new_path).grid(row = 3, column =0, columnspan = 1, pady=0)
        Button(self, text = "Browse", command = lambda: loadDir(new_path), width = 10).grid(row = 3, column = 1, pady=0)

        report_name_entry = StringVar()
        Label(self, text = 'Name for this report:').grid(row = 4, column = 0, pady=0)
        Entry(self, width=40,textvariable = report_name_entry).grid(row = 5, column =0, columnspan = 1, pady=0)

        Button(self, text = '  Create Report  ', command = lambda: multCommands(old_path, new_path, report_name_entry)).grid(row = 6, column = 0, pady=20)

    def centerWindow(self):
        w = 700
        h = 200
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()
        x = (sw - w)/2
        y = (sh - h)/2
        self.parent.geometry('%dx%d+%d+%d' % (w,h, x, y))


    def loadtemplate(self,entry):
        filename = tkFileDialog.askopenfilename()
        entry.set(filename)
#-------------------------------------------------------------------------------

def remove_trailing_slash(s):
    if s.endswith('/'):
        s = s[:-1]
    if s.endswith('\\'):
        s = s[:-1]
    return s

#-------------------------------------------------------------------------------

def convertSize(size):
    size_name = ("B", "KB", "MB", "GB", "TB")
    i = int(math.floor(math.log(size,1024)))
    p = math.pow(1024,i)
    s = round(size/p,2)
    if (s > 0):
        s = math.ceil(s)
        if i == 0:
            return "1 KB"
        else:
            return '%s %s' % (int(s),size_name[i])
    else:
        return '0 B'

#-------------------------------------------------------------------------------
def gatherFileSizes(product_location):
# Takes in one path and returns "row" which is ['FRE.tar.gz', '65 MB', 67912682L]
# file_name, human readable file size, byte size

    file_list = []
    i = 0
    for (path, dirs, files) in os.walk(product_location):
        print "path", path
        print "dirs", dirs
        print "files", files
        i += 1
        if i >= 1:
            break

    print "Number of files:", len(files)
    for f in files:
        file = product_location+'\\'+f
        b = os.path.getsize(file)
        accurate_kb = convertSize(b)
        row = [f, accurate_kb, b]
        file_list.append(row)

    return file_list

#-------------------------------------------------------------------------------
def prepareComparisonMatrix(paths_to_compare):
# paths_to_compare should always contain two paths at a time

    files_in_common = set()
    unmatched_files = set()
    old_file_names = set()
    new_file_names = set()

    old_matrix = {}
    new_matrix = {}
    old_product_location = paths_to_compare[0]
    new_product_location = paths_to_compare[1]

    old_file_list = gatherFileSizes(old_product_location)
    for old_row in old_file_list:
        old_row.append(old_product_location)
        file_name = old_row[0]
        old_file_names.add(file_name)
        old_matrix[file_name] = old_row
    new_file_list = gatherFileSizes(new_product_location)
    for new_row in new_file_list:
        new_row.append(new_product_location)
        file_name = new_row[0]
        new_file_names.add(file_name)
        new_matrix[file_name] = new_row

    unmatched_files = old_file_names ^ new_file_names
    files_in_common = old_file_names & new_file_names
    unmatched_matrix = prepareUnmatchedMatrix(old_matrix, new_matrix, unmatched_files)
    compareDirContents(old_matrix, new_matrix, files_in_common, unmatched_matrix)

#-------------------------------------------------------------------------------
def prepareUnmatchedMatrix(old_matrix, new_matrix, unmatched_files):
    # build a special table just for these unmatched file names
    unmatched_matrix = []
    header_row = ["File Name", "File Size", "Size in Bytes"]
    unmatched_files_array = list(unmatched_files)
    unmatched_files_array.sort()
    unmatched_files_array.sort(key=len, reverse=True)

    for u in unmatched_files_array:
        try:
            r = old_matrix[u]
        except:
            r = new_matrix[u]
        row = [r[0], r[1], format(r[2], ',')]
        unmatched_matrix.append(row)
    return unmatched_matrix

#-------------------------------------------------------------------------------
def compareDirContents(old_matrix, new_matrix, files_in_common, unmatched_matrix):
    dir_matrix = []
    header_row = ["File Name", "Old File Size", "Old Bytes", "New File Size", "New Bytes", "Difference", "% Diff", "Comment"]

    # Using old_matrix to iterate
    for k in files_in_common:
        file_name = k
        old_value_array = old_matrix[k]
        old_readable_size = old_value_array[1]
        byte_size_old = old_value_array[2]
        try:
            new_value_array = new_matrix[k]
            new_readable_size = new_value_array[1]
            byte_size_new = new_value_array[2]
        except:
            byte_size_new = int(0)
        diff = byte_size_new - byte_size_old
        percent_diff = return_percent_diff(byte_size_old, byte_size_new)
        comments = returnComments(diff)
        row = [file_name, old_readable_size, format(byte_size_old, ','), new_readable_size, format(byte_size_new, ','), format(diff, ','), percent_diff, comments]
        dir_matrix.append(row)
    makeReport(dir_matrix, unmatched_matrix)

#-------------------------------------------------------------------------------
def ThrowError(title, message, path):
    root = Tk()
    #root.title("Error")
    root.title(title)
    w = 900
    h = 200
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w)/2
    y = (sh - h)/2
    root.geometry('%dx%d+%d+%d' % (w,h, x, y))
    m = message
    m += '\n'
    m += path
    w = Label(root, text=m, width=120, height=10)
    w.pack()
    b = Button(root, text="OK", command=root.destroy, width=10)
    b.pack()
    mainloop()

#-------------------------------------------------------------------------------
def makeReport(dir_matrix, unmatched_matrix):

    # Works on one product array (worksheet) at a time
    global paths_to_compare
    product_location = paths_to_compare[0]
    product_location2 = paths_to_compare[1]
    placeholder_row = 0
    unmatched_starting_row = 0
    print "\ndir_matrix:"
    for d in dir_matrix:
        print d

    print "\nunmatched_matrix:"
    for d in unmatched_matrix:
        print d

    # Uses xlwt. xlsxwriter is better for writing to Excel documents
    workbook = xlwt.Workbook()
    borders = xlwt.Borders()
    alignmentR = xlwt.Alignment()
    alignmentL = xlwt.Alignment()
    alignmentC = xlwt.Alignment()

    font0 = xlwt.Font()
    font0.name = 'Arial'
    font0.bold = True

    font1 = xlwt.Font()
    font1.name = 'Arial'
    font1.colour_index = 2
    font1.bold = False

    bold_font = xlwt.Font()
    bold_font.name = 'Arial'
    bold_font.height = 320
    bold_font.bold = True

    """# May be: NO_LINE, THIN, MEDIUM, DASHED, DOTTED, THICK, DOUBLE, HAIR, MEDIUM_DASHED,
    THIN_DASH_DOTTED, MEDIUM_DASH_DOTTED, THIN_DASH_DOT_DOTTED, MEDIUM_DASH_DOT_DOTTED,
    SLANTED_MEDIUM_DASH_DOTTED, or 0x00 through 0x0D."""

    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left_colour = 0x40
    borders.right_colour = 0x40
    borders.top_colour = 0x40
    borders.bottom_colour = 0x40

    style1 = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # gray25, light_turquoise, pale_blue, sky_blue
    pattern.pattern_fore_colour = xlwt.Style.colour_map['tan']
    style1.pattern = pattern
    style1.borders = borders
    style1.font = font0
    alignmentL.horz = xlwt.Alignment.HORZ_LEFT
    style1.alignment = alignmentL

    style2 = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # gray25, light_turquoise, pale_blue, sky_blue
    pattern.pattern_fore_colour = xlwt.Style.colour_map['light_turquoise']
    style2.pattern = pattern
    style2.borders = borders
    style2.font = font0
    alignmentR.horz = xlwt.Alignment.HORZ_RIGHT
    style2.alignment = alignmentR

    style = xlwt.XFStyle()
    style.borders = borders
    alignmentR.horz = xlwt.Alignment.HORZ_RIGHT
    style.alignment = alignmentR

    style3 = xlwt.XFStyle()
    style3.borders = borders
    alignmentL.horz = xlwt.Alignment.HORZ_LEFT
    style3.alignment = alignmentL

    styleRED = xlwt.XFStyle()
    styleRED.borders = borders
    alignmentR.horz = xlwt.Alignment.HORZ_RIGHT
    styleRED.alignment = alignmentR
    styleRED.font = font1

    table_label_style = xlwt.XFStyle()
    table_label_style.font = bold_font

    # Create the worksheet
    worksheet_names = ["File Size Comparison"]

    for w in worksheet_names:
        workbook.add_sheet(w)

    worksheet_index = 0
    worksheet = workbook.get_sheet(worksheet_index)

    # --------------------------
    # Bold Labels

    starting_row = 1
    starting_column = 1

    worksheet.col(starting_column-1).width = 650                    # 3333 = 1" (one inch)
    worksheet.col(starting_column+0).width = 13332                  # 3333 = 1" (one inch)
    worksheet.col(starting_column+1).width = 3333                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+2).width = 3333                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+3).width = 3333                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+4).width = 3333                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+5).width = 3333                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+6).width = 2000                   # 3333 = 1" (one inch)
    worksheet.col(starting_column+7).width = 2850                   # 3333 = 1" (one inch)

    # --------------------------
    # Header row

    if dir_matrix:

        worksheet.write(starting_row, 1, 'Directory Contents Comparison', table_label_style)

        #['File Name', 'Old File Size', 'Old Bytes', 'New File Size', 'New Bytes', 'Difference', '% Diff', 'Comment']

        worksheet.write(starting_row+1, starting_column+0, 'File Name', style1)
        worksheet.write(starting_row+1, starting_column+1, "Old File Size", style1)
        worksheet.write(starting_row+1, starting_column+2, 'Old Bytes', style1)
        worksheet.write(starting_row+1, starting_column+3, 'New File Size', style1)
        worksheet.write(starting_row+1, starting_column+4, 'New Bytes', style1)
        worksheet.write(starting_row+1, starting_column+5, 'Difference', style1)
        worksheet.write(starting_row+1, starting_column+6, '% Diff', style1)
        worksheet.write(starting_row+1, starting_column+7, 'Comments', style1)

        # Start data row
        starting_row += 2
        # Skip the header row in dir_matrix

        dir_matrix = sorted(dir_matrix)
        f = 0
        ldm = len(dir_matrix)
        while f < ldm:
            line = dir_matrix[f]
            table_width = len(line)
            y = 0
            while y < table_width:
                r = f+starting_row
                c = y+starting_column
                k = line[y]
                s = style
                if k == "Drop":
                    s = styleRED
                if y == 0:
                    s = style3
                worksheet.write(r, c, k, s)
                unmatched_starting_row = r
                y += 1
            f += 1
        unmatched_starting_row += 1

    if unmatched_matrix:
        unmatched_starting_row += 1
        worksheet.write(unmatched_starting_row, 1, 'Unmatched File Names', table_label_style)
        unmatched_starting_row += 1

        worksheet.write(unmatched_starting_row, starting_column+0, 'File Name', style1)
        worksheet.write(unmatched_starting_row, starting_column+1, "File Size", style1)
        worksheet.write(unmatched_starting_row, starting_column+2, 'Size in Bytes', style1)
        unmatched_starting_row += 1

        # Skip the header row in dir_matrix
        unmatched_matrix.sort(key=len, reverse=True)

        f = 0
        ldm = len(unmatched_matrix)
        while f < ldm:
            line = unmatched_matrix[f]
            print line
            table_width = len(line)
            y = 0
            while y < table_width:
                r = f+unmatched_starting_row
                c = y+starting_column
                k = line[y]
                s = style
                if y == 0:
                    s = style3
                worksheet.write(r, c, k, s)
                placeholder_row = r
                y += 1
            f += 1
        placeholder_row += 1

    # --------------------------
    # Directory paths row

    if placeholder_row == 0:
        placeholder_row = unmatched_starting_row

    placeholder_row += 1

    path1 = "Path 1: "+product_location
    path2 = "Path 2: "+product_location2
    worksheet.write(placeholder_row, 1, 'Directories Compared', table_label_style)
    worksheet.write(placeholder_row+1, 1, path1)
    worksheet.write(placeholder_row+2, 1, path2)

    product_location = ""
    product_location2 = ""
    # --------------------------
    date = getDate()
    global report_folder
    global report_name
    document_name = report_name+'_'+date+".xls"
    save_as = report_folder+'\\'+document_name
    try:
        if not os.path.exists(report_folder):
            os.makedirs(report_folder)
    except:
        print "Folder creation error. System could not create folder here:", report_folder
        ThrowError("Error", "Folder creation error. System could not create folder here:", report_folder)
        sys.exit()
    try:
        workbook.save(save_as)
        print "\nProcess completed. Your report can be found here:", save_as
        ThrowError("Success!", "Process completed. Your report can be found here:", save_as)
    except:
        print "Write Error. Permission denied. Can't open:", save_as
        ThrowError("Error", "Write Error. Permission denied. Can't open:", save_as)
        sys.exit()


#-------------------------------------------------------------------------------
def get_timestamp():
    t = time.localtime()
    und = '_'
    date = getDate()
    timestamp = date+und+str(t[3])+und+str(t[4])+und+str(t[5])
    return timestamp

#-------------------------------------------------------------------------------
def getDate():
    now = datetime.datetime.now()
    #print now.month, now.day, now.year
    m = now.month
    d = now.day
    y = str(now.year)
    if m < 10:
        m = "0"+str(m)
    else:
        m = str(m)
    if d < 10:
        d = "0"+str(d)
    else:
        d = str(d)
    y = y[2:]

    formatted_date = m+d+y
    return formatted_date

#-------------------------------------------------------------------------------
def return_percent_diff(first_number, second_number):
        a = first_number
        b = second_number
        if a != b and a != 0:
            a = float(a)
            b = float(b)
            diff = a - b
            percent_diff = 100 * (b - a) / a
            if percent_diff < 1 and percent_diff > 0:
                percent_diff = "<1%"
            elif percent_diff < 0 and percent_diff > -1:
                percent_diff = "<1%"
            else:
                percent_diff = int(percent_diff)
                percent_diff = str(percent_diff)+"%"
        else:
            if a == 0 and b != 0:
                percent_diff = "100%"
            else:
                percent_diff = "n/a"
        return percent_diff

#-------------------------------------------------------------------------------
def returnComments(arg):
    dif = arg
    if dif == 0:
        comment = "n/a"
    elif dif > 0:
        comment = "Increase"
    elif dif < 0:
        comment = "Drop"
    return comment

#-------------------------------------------------------------------------------
def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

#-------------------------------------------------------------------------------
def setupEnvironment():
    script_dir = getScriptPath()
    global report_folder
    report_folder = script_dir+'\\'+'reports'

    try:
        if not os.path.exists(report_folder):
            os.makedirs(report_folder)
    except:
        print "Folder creation error. System could not create folder here:", report_folder
        ThrowError("Error", "Folder creation error. System could not create folder here:", report_folder)
        sys.exit()

#-------------------------------------------------------------------------------

def main():

    root = Tk()
    root.resizable(0, 0)
    Tk.wantobjects = 0
    app = pathDialog(root)
    root.mainloop()

    setupEnvironment()
    prepareComparisonMatrix(paths_to_compare)

    end_time_seconds = time.time()-start_time
    end_time_min = end_time_seconds/60
    print 'Execution time:', int(round(end_time_seconds)), 'seconds, or', int(end_time_min), "mins."


#-------------------------------------------------------------------------------
if __name__ == '__main__':
    main()
