import tkinter as tk
from tkinter import messagebox, filedialog
import pyodbc
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, NamedStyle
import sys, os, re, subprocess
import traceback
from io import StringIO
from statistics import mean
import xlrd
import simplekml
import datetime

if getattr(sys, 'frozen', False):
    CWD = sys._MEIPASS
else:
    CWD = os.path.dirname(os.path.realpath(__file__))

QC_PATH = None
TEMPLATE_FILE = os.path.join(CWD, 'template.xlsx')
FWD_TEST_LIST_FILE = None
PROJECT_NAME = ''
#DATE_RE = '_(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])[0-9]{2}\.mdb$'
MDB_RE = '\.mdb$'
PYTHON_DIR = ['C:\\Python27', 'C:\\ESRI\\Python2.7', 'C:\\Python2.7', 'C:\\ESRI\\Python27']
FONT_CHECKMARK = Font(name='Wingdings 2', size=11, bold=True)
ARCGIS_PY_SCRIPT = os.path.join(CWD, 'gen_arcgis_file.py')
MAKE_ARCGIS_FILE = False
MDB_FILES = None

Summary_Stats_List = []
Section_No_Dict = {}
Data_List = []
Station_IDs_Dict = {}
Insufficient_Tests_Dict = {}
Data_Headers = None
Fwd_Test_List = []

tk_root = None
tk_chkbox = None
tk_dir_entry = None
tk_dir_entry_str = None
tk_file_entry = None
tk_file_entry_str = None
tk_dir_button = None
tk_file_button = None
tk_run_button = None

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
        self.sect_no_check = ''
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
    for path in PYTHON_DIR:
        if os.path.isdir(path):
            for _path, subFolders, files in os.walk(path):
                for file_name in files:
                    m = re.match(r, file_name)
                    if m:
                        return os.path.join(_path, file_name)

def find_mdb_files(path):
    files_dict = {}
    for _path, subFolders, files in os.walk(path):
        for file_name in files:
            if is_mdb_file(file_name.lower()):
                file_path = os.path.join(_path, file_name)
                files_dict[file_name] = ((file_name, file_path))
    return sorted(files_dict.values())

def add_to_section_dict(mdb_data, mdb_file_name):
    global Section_No_Dict
    section_no_dict = {}
    stationID_col = get_col(Data_Headers, 'StationID')
    sect_col = get_col(Data_Headers, 'SECT_NO')
    if sect_col and stationID_col:
        for row in mdb_data:
            sect_no = row[sect_col]
            key = (row[stationID_col], sect_no)
            if key not in section_no_dict:
                if sect_no in Section_No_Dict:
                    count, mdb_file_names_dict = Section_No_Dict[sect_no]
                    count += 1
                    mdb_file_names_dict[mdb_file_name] = None
                    Section_No_Dict[sect_no] = [count, mdb_file_names_dict]
                else:
                    Section_No_Dict[sect_no] = [1, {mdb_file_name: None}]
                section_no_dict[key] = None

def add_sect_no_check(mdb_file_name, sect_no):
    for ss in Summary_Stats_List:
        if ss.mdb_file_name == mdb_file_name:
            if ss.sect_no_check:
                ss.sect_no_check += ', ' + sect_no
            else:
                ss.sect_no_check = 'Sect_No: ' + sect_no
            break

def read_fwd_test_list():
    global Fwd_Test_List
    wb = xlrd.open_workbook(FWD_TEST_LIST_FILE)
    ws = wb.sheet_by_index(0)
    if 'test list' in wb.sheet_names():
        ws = wb.sheet_by_name('test list')

    for i, row in enumerate(range(ws.nrows)):
        r = []
        for j, col in enumerate(range(ws.ncols)):
            r.append(ws.cell_value(i, j))
        Fwd_Test_List.append(r)
    if Fwd_Test_List:
        Fwd_Test_List[0].extend(['Field Tests', 'Compare Test Count', 'Insufficient Field Tests', 'Comments'])

def get_col(headers, header_name):
    try:
        fields = [h.upper() for h in headers]
        return fields.index(header_name.upper())
    except:
        return None

def get_col_no(headers, name_list):
    for name in name_list:
        index = get_col(headers, name)
        if index:
            return index

def style_cell(ws, row, col):
    c = ws.cell(row=row, column=col)
    c.font = FONT_CHECKMARK
    c.alignment = Alignment(horizontal='center')

def write_test_list_ws(wb):
    global Insufficient_Tests_Dict
    test_list_ws = wb.worksheets[1]
    if 'test list' in wb:
        test_list_ws = wb['test list']

    # write field test numbers, test list sheet
    headers = Fwd_Test_List[0]
    sect_no_col = get_col_no(headers, ['RT_NO', 'SECT_NO', 'FWD_NO']) or 0
    total_tests_col = get_col_no(headers, ['TOTAL TESTS','TOTAL', 'TOT_TEST']) or (len(headers) - 1)

    for col, header in enumerate(headers, start=1):
        test_list_ws.cell(row=1, column=col).value = header
    for i,row in enumerate(Fwd_Test_List[1:]):
        style = None
        sect_no = str(row[sect_no_col])
        if is_number(sect_no):
            sect_no = str(int(float(sect_no)))
        if sect_no in Section_No_Dict:
            field_tests, mdb_names_dict = Section_No_Dict[sect_no]
            reqd_tests = row[total_tests_col]
            row.append(field_tests)
            if is_number(reqd_tests):
                reqd_tests = int(round(float(reqd_tests), 0))
                check = reqd_tests - field_tests
                row.append(check)
                if check > 0:
                    for mdb_file_name in mdb_names_dict.keys():
                        add_sect_no_check(mdb_file_name, sect_no)
                    row.append('O')
                    style = [i + 2, len(row)]
                    Insufficient_Tests_Dict[sect_no] = None
        else:
            row.append('0')
            reqd_tests = row[total_tests_col]
            if is_number(reqd_tests):
                reqd_tests = int(round(float(reqd_tests), 0))
                row.append(reqd_tests)
                if reqd_tests > 0:
                    row.append('O')
                    style = [i + 2, len(row)]
                    Insufficient_Tests_Dict[sect_no] = None
        test_list_ws.append(row)
        if style:
            style_cell(test_list_ws, style[0], style[1])

def write_summary_ws(wb):
    summary_ws = wb.worksheets[0]
    if 'summary' in wb:
        summary_ws = wb['summary']

    # write summary stats
    for i,s in enumerate(Summary_Stats_List):
        if not s.sect_no_check and FWD_TEST_LIST_FILE:
            s.sect_no_check = 'P'
        row = (s.date, s.from_time, s.to_time, s.mdb_file_name, '', '', '', s.completed_tests, '', s.sect_no_check, s.max_station, '',
            s.min_surface_temp, s.max_surface_temp, '', s.min_air_temp, s.max_air_temp, '', s.station_ids or 'P')
        summary_ws.append(row)
        if not s.station_ids:
            style_cell(summary_ws, i + 2, 19)
        if not s.sect_no_check and FWD_TEST_LIST_FILE:
            style_cell(summary_ws, i + 2, 10)

def in_station_id_dict(row):
    drop_col = get_col(Data_Headers, 'DropID')
    file_col = get_col(Data_Headers, 'File')
    stationID_col = get_col(Data_Headers, 'StationID')
    if stationID_col and drop_col:
        key = (row[stationID_col], row[file_col], row[drop_col])
        if key in Station_IDs_Dict:
            return Station_IDs_Dict[key]
    return False

def write_station_ws(wb):
    # write station data
    if not 'Stations' in wb:
        wb.create_sheet('Stations')
    stations_ws = wb['Stations']

    highlight = NamedStyle(name="highlight")
    highlight.fill = PatternFill("solid", fgColor="f5b7b1")

    for col, header in enumerate(Data_Headers, start=1):
        stations_ws.cell(row=1, column=col).value = header

    sect_col = get_col(Data_Headers, 'SECT_NO')
    for i,row in enumerate(Data_List):
        checks = ['', '']
        j = in_station_id_dict(row)
        if j:
            # increasing deflection
            checks[0] = 'O'
        if sect_col and row[sect_col] in Insufficient_Tests_Dict:
            checks[1] = 'O'
        row.extend(checks)
        stations_ws.append(row)
        if j:
            # increasing deflection, highlight this cell
            xls_cell = stations_ws.cell(row=i + 2, column=j + 1)
            xls_cell.style = highlight
            style_cell(stations_ws, i + 2, len(row) - 1)
        if checks[1]:
            style_cell(stations_ws, i + 2, len(row))

def write_excel_file():
    wb = load_workbook(TEMPLATE_FILE)
    if Fwd_Test_List:
        write_test_list_ws(wb)
    write_summary_ws(wb)
    write_station_ws(wb)
    qc_file = os.path.join(QC_PATH, PROJECT_NAME + '.xlsx')
    wb.save(qc_file)
    return qc_file

def write_kml(file_path, color, row_dict=None):
    kml = simplekml.Kml()
    sect_col = get_col(Data_Headers, 'SECT_NO')
    slab_col = get_col(Data_Headers, 'SlabID')
    lat_col = get_col(Data_Headers, 'Latitude')
    long_col = get_col(Data_Headers, 'Longitude')
    if not (lat_col and long_col):
        return

    for i,row in enumerate(Data_List):
        if not row_dict or  i in row_dict:
            if sect_col and slab_col:
                pt = kml.newpoint(name=str(row[slab_col]), description=str(row[sect_col]), coords=[(row[long_col], row[lat_col])])
                pt.style.iconstyle.icon.href = 'http://www.google.com/intl/en_us/mapfiles/ms/icons/' + color + '-dot.png'
        #'http://maps.google.com/mapfiles/kml/paddle/' + color + '-circle.png'
    kml.save(file_path)

def write_kml_file():
    file_path = os.path.join(QC_PATH, PROJECT_NAME + '.kml')
    write_kml(file_path, 'blue')
    return file_path

def write_bad_drops_kml():
    if not Station_IDs_Dict:
        return
    row_dict = {}
    for i,row in enumerate(Data_List):
        if in_station_id_dict(row):
            row_dict[i] = None
    if row_dict:
        write_kml(os.path.join(QC_PATH, PROJECT_NAME + '_stations_with_increasing_deflection.kml'), 'red', row_dict)

def write_bad_sections_kml():
    sect_col = get_col(Data_Headers, 'SECT_NO')
    stationID_col = get_col(Data_Headers, 'StationID')
    if not (Insufficient_Tests_Dict and sect_col and stationID_col):
        return
    row_dict = {}
    section_no_dict = {}
    file_col = get_col(Data_Headers, 'File')
    
    for i,row in enumerate(Data_List):
        key = (row[sect_col], row[stationID_col], row[file_col])
        if key not in section_no_dict and row[sect_col] in Insufficient_Tests_Dict:
            row_dict[i] = None
            section_no_dict[key] = None
    if row_dict:
        write_kml(os.path.join(QC_PATH, PROJECT_NAME + '_sections_with_insufficient_tests.kml'), 'red', row_dict)

def check_coords(lats,longs):
        avg_lat = mean(lats)
        avg_long = mean(longs)
        '''
        for lat in lats:
            if lat <= avg_lat + 2 or lat >= avg_lat - 2:'''

def process_mdb_data(mdb_file_name, data_rows, data_headers):
    global Data_Headers, Data_List
    if not Data_Headers:
        Data_Headers = ['File'] + [col[0] for col in data_headers] + ['Decreasing Deflections', 'Insufficient Field Tests']
        slab_col = get_col(Data_Headers, 'SlabID')
        if slab_col:
            Data_Headers.insert(slab_col, 'SECT_NO')

    sect_col = get_col(Data_Headers, 'SECT_NO')
    mdb_data = []
    for row in data_rows:
        row = [mdb_file_name] + [x for x in row]
        if sect_col:
            row.insert(sect_col, left(row[sect_col], '-'))
        mdb_data.append(row)
        Data_List.append(row)

    data_transposed = [*zip(*mdb_data)]
    #summary stats
    completed_tests = len(mdb_data)
    station_col = get_col(Data_Headers, 'Station')
    if station_col:
        max_station = round(max(data_transposed[station_col]), 1)
    else:
        max_station = ''
    temp_col = get_col(Data_Headers, 'Surface')
    if temp_col:
        min_surface_temp = round(min(data_transposed[temp_col]), 1)
        max_surface_temp = round(max(data_transposed[temp_col]), 1)
    else:
        min_surface_temp = ''
        max_surface_temp = ''
    temp_col = get_col(Data_Headers, 'Air')
    if temp_col:
        min_air_temp = round(min(data_transposed[temp_col]), 1)
        max_air_temp = round(max(data_transposed[temp_col]), 1)
    else:
        min_air_temp = ''
        max_air_temp = ''
    time_col = get_col(Data_Headers, 'Time')
    if time_col:
        times = data_transposed[time_col]
        date = times[0].strftime('%m/%d/%Y')
        from_time = min(times).strftime('%H:%M')
        to_time = max(times).strftime('%H:%M')
    else:
        date = ''
        from_time = ''
        to_time = ''

    station_ids = {}
    d1_col = get_col(Data_Headers, 'D1')
    drop_col = get_col(Data_Headers, 'DropID')
    stationID_col = get_col(Data_Headers, 'StationID')

    if d1_col and stationID_col and drop_col:
        for row in mdb_data:
            for i in range(d1_col, d1_col + 7):
                if row[i] < row[i + 1]:
                    station_ids[str(row[stationID_col])] = None
                    key = (row[stationID_col], mdb_file_name, row[drop_col])
                    Station_IDs_Dict[key] = i + 1
                    break
    if station_ids:
        ids = station_ids.keys()
        station_ids = 'StationID ' + ', '.join(ids)
    else:
        station_ids = ''
    summary_stats = Summary_Stats(mdb_file_name, date, completed_tests, max_station, min_surface_temp, max_surface_temp, 
        min_air_temp, max_air_temp, station_ids, from_time, to_time)

    global Summary_Stats_List
    Summary_Stats_List.append(summary_stats)
    add_to_section_dict(mdb_data, mdb_file_name)

def query_mdb_data():
    DRV = '{Microsoft Access Driver (*.mdb)}'
    PWD = 'pw'
    for mdb in MDB_FILES:
        mdb_file_name = mdb[0]
        mdb_file_path = mdb[1]
        
        # connect to db
        connection = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, mdb_file_path, PWD))
        cursor = connection.cursor()

        data_query = 'SELECT * FROM Stations Inner Join Drops On Stations.StationID = Drops.StationID'
        data_rows = cursor.execute(data_query).fetchall()
        data_headers = cursor.description
        
        cursor.close()
        connection.close()
        print(mdb_file_name)
        process_mdb_data(mdb_file_name, data_rows, data_headers)

def check_selected_dir(wd):
    if wd:
        wd = wd.replace('/', '\\')
        if os.path.isdir(wd):
            mdb_files = find_mdb_files(wd)
            if mdb_files:
                return wd
    return False

def set_selected_dir():
    global tk_dir_entry_str
    options = {'title':'Select the QC Project Folder'}
    wd = filedialog.askdirectory(**options)
    if wd:
        wd = check_selected_dir(wd)
        if wd:
            tk_dir_entry_str.set(wd)
        else:
            messagebox.showinfo('', 'Please select the project folder containing ALL of the .mdb files.')

def check_selected_file(wf):
    if wf:
        wf = wf.replace('/', '\\')
        if os.path.isfile(wf) and \
        (wf.lower().endswith('.xls') or wf.lower().endswith('.xlsx') or wf.lower().endswith('.xlsm')):
            return wf
    return False

def set_selected_file():
    global tk_file_entry_str
    ftypes = [('Excel files', '*.xls;*.xlsx;*.xlsm;')]
    options = {'title':'Select the QC Specification Excel File', 'filetypes':ftypes}

    wd = tk_dir_entry_str.get()
    field_data = check_selected_dir(wd)
    if field_data:
        phase = left(field_data,'field_data')
        options['initialdir'] = phase
        if len(field_data) > len(phase):
            specification = os.path.join(phase,'specification')
            if os.path.isdir(specification):
                options['initialdir'] = specification
                fwd = os.path.join(specification, 'fwd')
                if os.path.isdir(fwd):
                    options['initialdir'] = fwd
                elif os.path.isdir(os.path.join(specification, 'fwd_setup')):
                    options['initialdir'] = os.path.join(specification, 'fwd_setup')

    wf = filedialog.askopenfilename(**options)
    if wf:
        wf = check_selected_file(wf)
        if wf:
            tk_file_entry_str.set(wf)
        else:
            messagebox.showinfo('', 'This program can only read Excel Test Specification files.')

def set_global_vars():
    global tk_root, QC_PATH, FWD_TEST_LIST_FILE, PROJECT_NAME, MAKE_ARCGIS_FILE, MDB_FILES
    error = False
    path = check_selected_dir(tk_dir_entry_str.get())
    test_list_file = tk_file_entry_str.get()
    if test_list_file:
        test_list_file = check_selected_file(test_list_file)
        if test_list_file == False:
            messagebox.showinfo('', 'This program can only read Excel Test Specification files.')
            error = True

    if not path:
        messagebox.showinfo('', 'Please select the project folder containing ALL of the .mdb files.')
        error = True

    if not error:
        QC_PATH = path
        if test_list_file:
            FWD_TEST_LIST_FILE = test_list_file
        MAKE_ARCGIS_FILE = bool(tk_chkbox.var.get())
        mdb_files = find_mdb_files(path)
        proj = mdb_files[0][0]
        PROJECT_NAME = left(proj, '.') + '_raw_qc_' + datetime.datetime.now().strftime('%Y%m%d')
        MDB_FILES = mdb_files

        tk_chkbox.configure(state='disable')
        tk_dir_entry.configure(state='readonly')
        tk_file_entry.configure(state='readonly')
        tk_dir_button.configure(state='disable')
        tk_file_button.configure(state='disable')
        tk_run_button.configure(state='disable')

        main()
        #tk_root.destroy()
        #tk_root.withdraw()

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

def main():
    if not (QC_PATH and MDB_FILES):
        exit()
    query_mdb_data()
    if FWD_TEST_LIST_FILE:
        read_fwd_test_list()
    qc_file = write_excel_file()
    kml_file = write_kml_file()
    write_bad_sections_kml()
    write_bad_drops_kml()
    make_arcgis_shape_file(kml_file)
    open_files(qc_file)
    print('Success!')

def set_up_gui():
    global tk_root, tk_chkbox, tk_dir_entry, tk_file_entry, tk_file_entry_str, tk_dir_entry_str, tk_dir_button, tk_file_button, tk_run_button
    tk_root = tk.Tk()
    tk_root.wm_title('Auto QC by Eric Rothfels ;)')
    tk.Label(tk_root, text='Select the Project Folder Containing .mdb files', 
        font = "Helvetica 18").grid(row=0, column=1, columnspan=3, padx=15,pady=15)
    tk.Label(tk_root, text='Select the FWD Test List File', 
        font = "Helvetica 18").grid(row=2, column=1, columnspan=3, padx=15,pady=15)
    v = tk.IntVar()
    tk_chkbox = tk.Checkbutton(tk_root, text="Make ArcGIS Shape File from FWD Tests", variable=v)
    tk_chkbox.var = v
    tk_chkbox.grid(row=4, column=2, columnspan=2)
    #tk_chkbox.select()
    tk.Label(tk_root, text='  \u2794 ', font = "Helvetica 14").grid(row=5, column=3, columnspan=1)
    tk_dir_button = tk.Button(tk_root, text='Browse', font='Helvetica 12', command=set_selected_dir)
    tk_dir_button.grid(row=1, column=6, sticky=tk.W, padx=5,pady=5)
    tk_file_button = tk.Button(tk_root, text='Browse', font='Helvetica 12', command=set_selected_file)
    tk_file_button.grid(row=3, column=6, sticky=tk.W, padx=5,pady=5)
    tk_run_button = tk.Button(tk_root, text='Run', font='Helvetica 14', command=set_global_vars)
    tk_run_button.grid(row=5, column=7, sticky=tk.W, padx=15,pady=15)
    
    tk.Label(tk_root, text='Project Folder:', 
        font = "Helvetica 12").grid(row=1, column=0, columnspan=1, padx=5,pady=5)
    tk.Label(tk_root, text='FWD Test List File: (optional)', 
        font = "Helvetica 12").grid(row=3, column=0, columnspan=1, padx=5,pady=5)
    tk_dir_entry_str = tk.StringVar()
    tk_dir_entry = tk.Entry(tk_root, textvariable=tk_dir_entry_str, width=100, readonlybackground='grey82')
    tk_dir_entry.grid(row=1, column=1, columnspan=5)
    tk_file_entry_str = tk.StringVar()
    tk_file_entry = tk.Entry(tk_root, textvariable=tk_file_entry_str, width=100, readonlybackground='grey82')
    tk_file_entry.grid(row=3, column=1, columnspan=5)
    #centre the window
    tk_root.eval('tk::PlaceWindow %s center' % tk_root.winfo_pathname(tk_root.winfo_id()))
    tk_root.mainloop()

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
    print("""\nWelcome to Auto QC \n\nClose this window at any time to Quit the Program.\n
    \n""")
    set_up_gui()
