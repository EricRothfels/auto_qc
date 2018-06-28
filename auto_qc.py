import tkinter as tk
from tkinter import messagebox, filedialog
import pyodbc
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import sys, os, re, subprocess
import traceback
from io import StringIO
from statistics import mean
import xlrd
import simplekml

if getattr(sys, 'frozen', False):
    CWD = sys._MEIPASS
else:
    CWD = os.path.dirname(os.path.realpath(__file__))

QC_PATH = None
TEMPLATE_FILE = os.path.join(CWD, 'template.xlsx')
DROPS_TEPLATE = os.path.join(CWD, 'drops_template.xlsx')
SPEC_FILE = None
PROJECT_NAME = ''

Summary_Stats_List = []
Section_No_Dict = {}
Station_Data_List = []
Section_No_List = []
Station_IDs_Dict = {}
Insufficient_Tests_Dict = {}
Drops_List = []

Drops_Headers = None

tk_root = None
tk_chkbox = None
#DATE_RE = '_(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])[0-9]{2}\.mdb$'
MDB_RE = '\.mdb$'
#GEP = 'C:\\Program Files\\Google\\Google Earth Pro\\client\\googleearth.exe'
PYTHON_DIR = 'C:\\ESRI\\Python2.7'
FONT_CHECKMARK = Font(name='Wingdings 2', size=11, bold=True)
ARCGIS_PY_SCRIPT = os.path.join(CWD, 'gen_arcgis_file.py')
MAKE_ARCGIS_FILE = False

class Summary_Stats:
    def __init__(self, mdb_file_name, date, completed_tests, max_station, min_surface_temp, max_surface_temp, 
        min_air_temp, max_air_temp, station_ids, from_time, to_time):
        self.date = date
        self.max_station = max_station
        self.min_surface_temp = min_surface_temp
        self.max_surface_temp = max_surface_temp
        self.min_air_temp = min_air_temp
        self.max_air_temp = max_air_temp
        self.station_ids = station_ids
        self.mdb_file_name = mdb_file_name
        self.completed_tests = completed_tests
        self.rt_no_check = ''
        self.from_time = from_time
        self.to_time = to_time

def is_number(value):
  try:
    float(value)
    return True
  except:
    return False

def left(string, delim):
    pos = string.rfind(delim)
    if pos == -1:
        return string
    return string[:pos]

def is_mdb_file(file):
    m = re.search(MDB_RE, file)
    if m:
        return True
    return False

def find_python27():
    r = 'python.exe'
    if os.path.isdir(PYTHON_DIR):
        for _path, subFolders, files in os.walk(PYTHON_DIR):
            for file_name in files:
                m = re.match(r, file_name)
                if m:
                    return os.path.join(_path, file_name)
    return False

def find_mdb_files(path):
    files_dict = {}
    for _path, subFolders, files in os.walk(path):
        for file_name in files:
            if is_mdb_file(file_name.lower()):
                file_path = os.path.join(_path, file_name)
                files_dict[file_name] = ((file_name, file_path))
    return sorted(files_dict.values())

def get_section_no_list(stations_rows):
    section_no_list = []
    for row in stations_rows:
        slab_id = row[8]
        section_no = left(slab_id, '-')
        section_no_list.append(section_no)
    return section_no_list

def add_to_dict(_dict, section_no_list, mdb_file_name):
    for section_no in section_no_list:
        if section_no in _dict:
            count, mdb_file_names_dict = _dict[section_no]
            count += 1
            mdb_file_names_dict[mdb_file_name] = None
            _dict[section_no] = [count, mdb_file_names_dict]
        else:
            _dict[section_no] = [1, {mdb_file_name: None}]
    return _dict

def find_header_index(ws, header_name):
    pos = len(ws.rows.next())
    for i in range(1, pos + 1):
        if ws.cell(row=1, column=i).value == header_name:
            return i
    return 0

def add_rt_no_check(mdb_file_name, rt_no):
    for ss in Summary_Stats_List:
        if ss.mdb_file_name == mdb_file_name:
            if ss.rt_no_check:
                ss.rt_no_check += ', ' + rt_no
            else:
                ss.rt_no_check = 'Section No: ' + rt_no
            break

def get_spec():
    wb = xlrd.open_workbook(SPEC_FILE)
    ws = wb.sheet_by_index(0)
    if 'test list' in wb.sheet_names():
        ws = wb.sheet_by_name('test list')

    rows = []
    for i, row in enumerate(range(ws.nrows)):
        r = []
        for j, col in enumerate(range(ws.ncols)):
            r.append(ws.cell_value(i, j))
        rows.append(r)
    return rows

def get_col(headers, header_name):
    try:
        return headers.index(header_name)
    except:
        return None

def get_col_no(headers, name_list):
    index = None
    fields = [h.upper() for h in headers]
    for name in name_list:
        index = get_col(fields, name.upper())
        if index:
            break
    return index

def style_cell(ws, row, col):
    c = ws.cell(row=row, column=col)
    c.font = FONT_CHECKMARK
    c.alignment = Alignment(horizontal='center')

def write_test_list_ws(wb):
    global Insufficient_Tests_Dict
    spec_rows = get_spec()
    test_list_ws = wb.worksheets[1]
    if 'test list' in wb:
        test_list_ws = wb['test list']

    # write field test numbers, test list sheet
    headers = spec_rows[0]
    rt_no_col = get_col_no(headers, ['RT_NO', 'SECT_NO', 'FWD_NO']) or 0
    total_tests_col = get_col_no(headers, ['TOTAL TESTS','TOTAL', 'TOT_TEST']) or (len(headers) - 1)
    headers.extend(['Field Tests', 'Compare Test Count', 'Insufficient Field Tests', 'Comments'])
    for col, header in enumerate(headers, start=1):
        test_list_ws.cell(row=1, column=col).value = header
    for i,row in enumerate(spec_rows[1:]):
        style = None
        rt_no = str(row[rt_no_col])
        if is_number(rt_no):
            rt_no = str(int(float(rt_no)))
        if rt_no in Section_No_Dict:
            field_tests, mdb_names_dict = Section_No_Dict[rt_no]
            reqd_tests = row[total_tests_col]
            row.append(field_tests)
            if is_number(reqd_tests):
                reqd_tests = int(round(float(reqd_tests), 0))
                check = reqd_tests - field_tests
                row.append(check)
                if check > 0:
                    for mdb_file_name in mdb_names_dict.keys():
                        add_rt_no_check(mdb_file_name, rt_no)
                    row.append('O')
                    style = [i + 2, len(row)]
                    Insufficient_Tests_Dict[rt_no] = None
        else:
            row.append('0')
            reqd_tests = row[total_tests_col]
            if is_number(reqd_tests):
                reqd_tests = int(round(float(reqd_tests), 0))
                row.append(reqd_tests)
                if reqd_tests > 0:
                    row.append('O')
                    style = [i + 2, len(row)]
                    Insufficient_Tests_Dict[rt_no] = None
        test_list_ws.append(row)
        if style:
            style_cell(test_list_ws, style[0], style[1])

def write_summary_ws(wb):
    summary_ws = wb.worksheets[0]
    if 'summary' in wb:
        summary_ws = wb['summary']

    # write summary stats
    for i,s in enumerate(Summary_Stats_List):
        row = (s.date, s.from_time, s.to_time, s.mdb_file_name, '', '', '', s.completed_tests, '', s.rt_no_check or 'P', s.max_station, '',
            s.min_surface_temp, s.max_surface_temp, '', s.min_air_temp, s.max_air_temp, '', s.station_ids or 'P')
        summary_ws.append(row)
        if not s.station_ids:
            style_cell(summary_ws, i + 2, 19)
        if not s.rt_no_check:
            style_cell(summary_ws, i + 2, 10)

def write_station_ws(wb):
    # write station data
    if not 'Stations' in wb:
        wb.create_sheet('Stations')
    stations_ws = wb['Stations']
    for i,row in enumerate(Station_Data_List):
        rt_no = Section_No_List[i]
        row.insert(9, rt_no)
        checks = ['', '']
        if str(row[2]) + row[0] in Station_IDs_Dict:
            checks[0] = 'O'
        if rt_no in Insufficient_Tests_Dict:
            checks[1] = 'O'
        row.extend(checks)
        stations_ws.append(row)
        if checks[0]:
            style_cell(stations_ws, i + 2, len(row) - 1)
        if checks[1]:
            style_cell(stations_ws, i + 2, len(row))

def write_excel_file():
    wb = load_workbook(TEMPLATE_FILE)
    write_test_list_ws(wb)
    write_summary_ws(wb)
    write_station_ws(wb)
    qc_file = os.path.join(QC_PATH, PROJECT_NAME + '_qc.xlsx')
    wb.save(qc_file)
    return qc_file

def write_drops_file():
    wb = load_workbook(DROPS_TEPLATE)
    drops_ws = wb.worksheets[0]
    for col, header in enumerate(Drops_Headers, start=1):
        drops_ws.cell(row=1, column=col).value = header
    for i,row in enumerate(Drops_List):
        drops_ws.append(row)
        if row[-1] is 'O':
            style_cell(drops_ws, i + 2, len(row))
        
    drops_file = os.path.join(QC_PATH, PROJECT_NAME + '_drops.xlsx')
    wb.save(drops_file)

def write_kml(file_path, data, color):
    kml = simplekml.Kml()
    for row in data:
        pt = kml.newpoint(name=str(row[0]), description=str(row[1]), coords=row[2])
        pt.style.iconstyle.icon.href = 'http://www.google.com/intl/en_us/mapfiles/ms/icons/' + color + '-dot.png'
        #'http://maps.google.com/mapfiles/kml/paddle/' + color + '-circle.png'
    kml.save(file_path)

def write_kml_file():
    data = []
    i = 0
    for row in Station_Data_List:
        data.append([row[9], Section_No_List[i], [(row[21], row[20])]])
        i += 1
    file_path = os.path.join(QC_PATH, PROJECT_NAME + '_qc.kml')
    write_kml(file_path, data, 'blue')
    return file_path

def write_bad_sections_kml():
    if not Insufficient_Tests_Dict:
        return
    #headers = ('Section No', 'Slab ID', 'Latitude', 'Longitude')
    data = []
    i = 0
    for row in Station_Data_List:
        rt_no = Section_No_List[i]
        if rt_no in Insufficient_Tests_Dict:
            data.append([row[9], rt_no, [(row[21], row[20])]])
        i += 1
    write_kml(os.path.join(QC_PATH, PROJECT_NAME + ' Sections with Insufficient Tests.kml'), data, 'red')

def process_mdb_data(mdb_file_name, stations_rows, drops_rows):
    stations_data = [*zip(*stations_rows)]
    #summary stats
    completed_tests = len(stations_rows)
    max_station = round(max(stations_data[2]), 1)
    min_surface_temp = round(min(stations_data[12]), 1)
    max_surface_temp = round(max(stations_data[12]), 1)
    min_air_temp = round(min(stations_data[13]), 1)
    max_air_temp = round(max(stations_data[13]), 1)

    time = stations_data[15]
    date = time[0].strftime('%m/%d/%Y')
    from_time = min(time).strftime('%H:%M')
    to_time = max(time).strftime('%H:%M')

    global Drops_List
    station_ids = {}
    for row in drops_rows:
        rowAsList = [mdb_file_name] + [x for x in row]
        for i in range(6,13):
            if rowAsList[i] < rowAsList[i + 1]:
                station_ids[str(rowAsList[1])] = None
                rowAsList.append('O')
                break
        Drops_List.append(rowAsList)

    if station_ids:
        ids = station_ids.keys()
        station_ids = 'Station ID ' + ', '.join(ids)
        for stn in ids:
            Station_IDs_Dict[stn + mdb_file_name] = stn
    else:
        station_ids = ''
    summary_stats = Summary_Stats(mdb_file_name, date, completed_tests, max_station, min_surface_temp, max_surface_temp, 
        min_air_temp, max_air_temp, station_ids, from_time, to_time)

    # whole project data
    global Section_No_Dict, Section_No_List, Summary_Stats_List, Station_Data_List
    Summary_Stats_List.append(summary_stats)
    section_no_list = get_section_no_list(stations_rows)
    Section_No_List.extend(section_no_list)
    add_to_dict(Section_No_Dict, section_no_list, mdb_file_name)
    for row in stations_rows:
        rowAsList = [mdb_file_name] + [x for x in row]
        Station_Data_List.append(rowAsList)

def check_coords(lats,longs):
        avg_lat = mean(lats)
        avg_long = mean(longs)
        '''
        for lat in lats:
            if lat <= avg_lat + 2 or lat >= avg_lat - 2:'''

def query_mdb_data(mdb_files):
    DRV = '{Microsoft Access Driver (*.mdb)}'
    PWD = 'pw'
    for mdb in mdb_files:
        mdb_file_name = mdb[0]
        mdb_file_path = mdb[1]
        
        # connect to db
        connection = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, mdb_file_path, PWD))
        cursor = connection.cursor()

        stations_query = 'SELECT * FROM Stations;'
        stations_rows = cursor.execute(stations_query).fetchall()
        stations_fields = [column[0] for column in cursor.description]

        drops_query = 'SELECT * FROM Drops;'
        drops_rows = cursor.execute(drops_query).fetchall()
        
        global Drops_Headers
        if not Drops_Headers:
            drops_fields = ['File'] + [column[0] for column in cursor.description]
            drops_fields.append('Decreasing Deflections')
            Drops_Headers = drops_fields

        cursor.close()
        connection.close()
        print(mdb_file_name)
        process_mdb_data(mdb_file_name, stations_rows, drops_rows)

def set_global_vars():
    global tk_root, QC_PATH, SPEC_FILE, PROJECT_NAME, MAKE_ARCGIS_FILE
    path = None
    spec_file = None
    options = {'title':'Select the QC Project Folder'}
    while True:
        wd = filedialog.askdirectory(**options)
        if not wd:
            break
        wd = wd.replace('/', '\\')
        if wd and os.path.isdir(wd):
            mdb_files = find_mdb_files(wd)
            if mdb_files:
                proj = mdb_files[0][0]
                PROJECT_NAME = left(proj, '.')
                path = wd
                ftypes = [('Excel files', '*.xls;*.xlsx;*.xlsm;')]
                options = {'title':'Select the QC Specification Excel File', 'filetypes':ftypes}
                wd = left(wd,'field_data\\fwd')
                if len(wd) < len(path):
                    fwd_dir = os.path.join(wd,'specification')
                    if os.path.isdir(fwd_dir):
                        wd = fwd_dir
                        fwd_dir = os.path.join(wd, 'fwd')
                        if os.path.isdir(fwd_dir):
                            wd = fwd_dir
                    options['initialdir'] = wd
                while True:
                    wf = filedialog.askopenfilename(**options)
                    if not wf:
                        break
                    wf = wf.replace('/', '\\')
                    if os.path.isfile(wf) and (wf.lower().endswith('.xls') or wf.lower().endswith('.xlsx') or wf.lower().endswith('.xlsm')):
                        spec_file = wf
                        break
                    else:
                        messagebox.showinfo('', 'This program can only read Excel Test Specification files.')
                break
            else:
                messagebox.showinfo('', 'Please select the project folder containing ALL of the .mdb files.')

    if path and os.path.isdir(path) and spec_file:
        QC_PATH = path
        SPEC_FILE = spec_file
        MAKE_ARCGIS_FILE = bool(tk_chkbox.var.get())
        tk_root.destroy()

def open_qc_file(qc_file):
    excel_p = subprocess.Popen('"' + qc_file + '"', shell=True)

def to_cmd_str(args):
    return ('"{}" '*len(args)).format(*args).rstrip()

def make_arcgis_shape_file(kml_file):
    global MAKE_ARCGIS_FILE
    if MAKE_ARCGIS_FILE:
        print('\nMaking ArcGIS layer with tests\n')
        py_exe = find_python27()
        if py_exe:
            cmd = to_cmd_str([py_exe, ARCGIS_PY_SCRIPT, kml_file, QC_PATH])
            try:
                ret = subprocess.check_call(cmd, shell=True)
            except Exception as e:
                handle_exception(e)

def open_files(qc_file):
    open_qc_file(qc_file)
    subprocess.Popen('explorer ' + QC_PATH)
    '''if os.path.isfile(GEP):
        os.startfile(GEP)'''

def main():
    print("""\nWelcome to Auto QC \n\nClose this window at any time to Quit the Program.\n
    \n""")
    #time.sleep(0.85)

    global tk_root, tk_chkbox
    tk_root = tk.Tk()
    tk_root.wm_title('Auto QC by Eric Rothfels ;)')
    tk.Label(tk_root, text='Select the Project Folder Containing .mdb files\nThen, Select the FWD Test List File', 
        font = "Helvetica 18").grid(row=0, column=1, columnspan=3, padx=15,pady=15)
    v = tk.IntVar()
    tk_chkbox = tk.Checkbutton(tk_root, text="Make ArcGIS Shape File from Tests", variable=v)
    tk_chkbox.var = v
    tk_chkbox.grid(row=4, column=2, columnspan=2)
    #tk_chkbox.select()
    tk.Label(tk_root, text='  \u2794 ', font = "Helvetica 14").grid(row=5, column=3, columnspan=1)
    tk.Button(tk_root, text='OK', font='Helvetica 14', command=set_global_vars).grid(row=5, column=4, sticky=tk.W, padx=15,pady=15)
    #centre the window
    tk_root.eval('tk::PlaceWindow %s center' % tk_root.winfo_pathname(tk_root.winfo_id()))
    tk_root.mainloop()

    root = tk.Tk()
    root.withdraw()

    if not QC_PATH:
        exit()
    mdb_files = find_mdb_files(QC_PATH)
    query_mdb_data(mdb_files) 
    qc_file = write_excel_file()
    write_drops_file()
    kml_file = write_kml_file()
    write_bad_sections_kml()
    make_arcgis_shape_file(kml_file)
    open_files(qc_file)
    print('Success!')

def get_exception_stack():
    old_stderr = sys.stderr
    sys.stderr = mystderr = StringIO()
    traceback.print_exc()
    sys.stderr = old_stderr
    return mystderr.getvalue()

def handle_exception(e):
    error = get_exception_stack()
    if DEV_MODE:
        print(error, file=sys.stderr)
    else:
        file = None
        #send_email(error, file)
    error += '\n\nPress OK to Continue, or Cancel to Quit the Program'
    if not messagebox.askokcancel('Error', error, icon=messagebox.ERROR):
        os._exit(0)

if __name__=="__main__":
    main()
