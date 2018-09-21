# Script to link PowerFactory loads to PowerGLF
# 19/09/2018 1.2 Peace Lekalakala
import datetime
import os
import powerfactory
import pandas as pd
import sqlite3
import sys
import uuid

from tkinter import *
from tkinter import filedialog
from tkinter import ttk


def GUI():

    def get_mv_shape():
        mv_shapefile.set(filedialog.askopenfilename(
                    defaultextension = ".shp",
                    filetypes = [('GIS Shapefile', '.shp')]))
                    
    def get_landuse_shape():
        glf_gis_shapefile.set(filedialog.askopenfilename(
                    defaultextension = ".shp",
                    filetypes = [('GIS Shapefile', '.shp')]))
    
    def get_glf_file():
        glf_file.set(filedialog.askopenfilename(
                    defaultextension = ".glf",
                    filetypes = [('GLF SQLite file', '.glf')]))
                    
    def display_choice(event=None):
        if 'Use SharePoint data' in select.get():
            # Clear previous
            canvas = Canvas(master, width=900, height=200)
            canvas.place(x=2, y=200)
                        
            l4 = Label(master, font="Bold", text="UserName:")
            l4.place(x=2, y=200)
            e4 = Entry(master, width=50, textvariable=user_name)
            e4.place(x=2, y=235)
    
            l5 = Label(master, font="Bold", text="Password:")
            l5.place(x=450, y=200)
            e5 = Entry(master, show="*", width=50)
            e5.place(x=450, y=235)
            
            link = Label(master, text="View SharePoint data", fg="blue", cursor="hand2")
            link.place(x=2, y=305)
            link.bind("<Button-1>", open_link)
        
        else:
            canvas = Canvas(master, width=900, height=200)
            canvas.place(x=2, y=200)
            
            l1 = Label(master, font="Bold", text="MV Stations Shapefile:")
            l1.place(x=2, y=200)
            e1 = Entry(master, width=100, textvariable=mv_shapefile)
            e1.place(x=3, y=235)
            b1 = Button(master, text="Browse...", command=get_mv_shape)
            b1.place(x=820, y=234)
    
            l2 = Label(master, font="Bold", text="GLF GIS Land Use Shapefile:")
            l2.place(x=2, y=270)
            e2 = Entry(master, width=100, textvariable=glf_gis_shapefile)
            e2.place(x=3, y=305)
            b2 = Button(master, text="Browse...", command=get_landuse_shape)
            b2.place(x=820, y=304)
            
    def open_link(event=None):
        import webbrowser
        webbrowser.open_new(r"https://mysites.eskom.co.za/personal/elec_lekalap/Planning/linkedzones/lz_table.aspx")
                    
    def set_values():
        input_data['MV Stations'] = e1.get()
        input_data['Load Zones'] = e2.get()
        input_data['GLF File'] = e3.get()
        input_data['User'] = e4.get()
        input_data['Password'] = e5.get()
        input_data['Method'] = select.get()
        input_data['Base Year'] = e6.get()
        
        # Exit gui
        master.destroy()
        
    def cancel_op():
        master.destroy()
        sys.exit()
    
        
       
    master = Tk()
    master.geometry("900x550")
    master.title("PowerGLF Linker")
        
    glf_gis_shapefile = StringVar()
    mv_shapefile = StringVar()
    glf_file = StringVar()
    user_name = StringVar()
    user_name.set(os.getenv('username'))
    year = StringVar()
    year.set(int(datetime.datetime.now().year) - 1)
    
      
    g_drive = r'\\ecrfnp01\USERG01\Sharedat\Planning_Central\Planning_Users\Peace\glfdata'
    glf_gis_shapefile.set(os.path.join(g_drive, 'GIS Load Zones.shp'))
    mv_shapefile.set(os.path.join(g_drive, 'GOU MV STATIONS_052017.shp'))
    
    
    open_note = "The script requires GIS shapefiles from PowerGLF checkout file"
    
    
    Label(master, text=open_note).place(x=2, y=1)
    
    ttk.Separator(master).place(x=1, y=140, relwidth=2)
    #ttk.Separator(master).place(x=1, y=320, relwidth=2)
    #ttk.Separator(master).place(x=1, y=400, relwidth=2)
    e4 = Entry(master, width=50, textvariable=user_name)
    e5 = Entry(master, show="*", width=50)
    
    Label(master, font="Bold", text="Select Checked out GLF file:").place(x=2, y=65)
    e3 = Entry(master, width=100, textvariable=glf_file)
    e3.place(x=3, y=95)
    b3 = Button(master, text="Browse...", command=get_glf_file)
    b3.place(x=820, y=94)
    
    Label(master, text="Select Linking Method:").place(x=2, y=150)
    select = ttk.Combobox(width=25,value=['Use SharePoint data', 'Use imported shapefiles'])
    select.set("Use imported shapefiles")
    select.place(x=180, y=149)  
    select.bind('<<ComboboxSelected>>', display_choice)

    Label(master, text="Enter Base Year:").place(x=450, y=150)
    e6 = Entry(master, width=20, textvariable=year)    
    e6.place(x=580, y=149)
    
    l1 = Label(master, font="Bold", text="MV Stations Shapefile:")
    l1.place(x=2, y=200)
    e1 = Entry(master, width=100, textvariable=mv_shapefile)
    e1.place(x=3, y=235)
    b1 = Button(master, text="Browse...", command=get_mv_shape)
    b1.place(x=820, y=234)
    
    l2 = Label(master, font="Bold", text="GLF GIS Land Use Shapefile:")
    l2.place(x=2, y=270)
    e2 = Entry(master, width=100, textvariable=glf_gis_shapefile)
    e2.place(x=3, y=305)
    b2 = Button(master, text="Browse...", command=get_landuse_shape)
    b2.place(x=820, y=304)   
    
    Button(master, font="Bold", text="Cancel", width=35, command=cancel_op).place(x=5, y=460)
    Button(master, font="Bold", text="Link Loads", width=35, command=set_values).place(x=450, y=460)
        
    master.mainloop( )
    

def get_object_ids(table, object):
    """convert binary to text"""
    
    object_ids = connection.execute("""SELECT
                 substr(hguid, 7, 2) || substr(hguid, 5, 2) || 
                 substr(hguid, 3, 2) || substr(hguid, 1, 2) || '-' ||
                 substr(hguid, 11, 2) || substr(hguid, 9, 2) || '-' || 
                 substr(hguid, 15, 2) || substr(hguid, 13, 2) || '-' || 
                 substr(hguid, 17, 4) || '-' || 
                 substr(hguid, 21, 12) AS {1}
                 FROM (SELECT hex({1}) AS hguid FROM {0})""".format(table,
                                    object))
                                    
    object_ids = [i[0] for i in object_ids.fetchall()]
    return object_ids

    
def get_pf_name(name):
    return str(name).replace('/', ' ').replace('-', ' ')
    
 
def get_linked_zones(gis_shapefile, mv_shapefile):
    import geopandas as gpd
    
    glf_df = gpd.read_file(gis_shapefile)
    glf_df = glf_df.to_crs("+proj=longlat +ellps=WGS84 +datum=WGS84 +no_defs")

    mv_df = gpd.read_file(mv_shapefile)
    mv_df = mv_df.to_crs({'init':'epsg:4326'})

    load_zones_names = connection.execute('select id from {}'.format(load_zones_table))
    load_zones_names = [i[0] for i in load_zones_names.fetchall()]

    zone_ids = get_object_ids(load_zones_table, 'fguid')
    load_zones_glf = dict(zip(load_zones_names, zone_ids))

    # Get corresponding load/transformers zone
    merged_df = gpd.sjoin(mv_df, glf_df, how="left", op="within")
    merged_df['PF Name'] = merged_df.standard_l.apply(lambda x: get_pf_name(x))

    identified_zones = dict(zip(merged_df['PF Name'], merged_df.lob_id))
    
    object_zone_ids = {}
    for k, v in identified_zones.items():
        try:
            object_zone_ids[k] = load_zones_glf[v]
        except KeyError:
            pass
    
    return object_zone_ids
    

def get_shapepoint_data(username, password):
    from requests_ntlm import HttpNtlmAuth
    import requests
        
    url = 'https://mysites.eskom.co.za/personal/elec_lekalap/Planning/linkedzones/lz_table.aspx'
    auth = HttpNtlmAuth('elec\\' + username, password)
    html = requests.get(url, auth=auth)
    html_data = html.text
    app.PrintPlain(html)
    try:
        df = pd.read_html(html_data)[0]
    except ValueError:
        csv = r'\\ecrfnp01\USERG01\Sharedat\Planning_Central\Planning_Users\Peace\glfdata\zones.csv'
        df = pd.read_csv(csv, encoding = "ISO-8859-1")
    df['PF Name'] = df.standard_l.apply(lambda x: get_pf_name(x))
    object_zone_ids = dict(zip(df['PF Name'], df.lob_fguid))
    
    load_names = [i.fold_id.loc_name for i in loads]
    object_zone_ids = {k: v for k, v in object_zone_ids.items() 
                       if k in load_names}
        
    
    return object_zone_ids
    
    
def get_zone_id(load_name):
    from fuzzywuzzy import fuzz
    
    name_ratios = [fuzz.ratio(load_name, j) 
                   for i, j in enumerate(object_zone_ids.keys())]
    max_value = max(name_ratios)
    index_ = [i for i, j in enumerate(name_ratios) if j == max_value][0]
    key = list(object_zone_ids.keys())[index_]
    return object_zone_ids[key]
    

def set_base_year(base_year):
    connection.execute('UPDATE tb_sys_properties SET baseyear = {}'.fomat(base_year))
            
    
    
def link_loads(loads):
    loads_data = []
        
    for load in loads:
        load_name = load.GetFullName()
        load_desc = load.loc_name
        load_capacity = 0 #load.slini
        load_short_name = load.fold_id.loc_name
        try:
            load_object_zone_id = object_zone_ids[load_short_name]
        except KeyError:
            load_object_zone_id = get_zone_id(load_short_name)
            
        row_data = (str(uuid.uuid1()), project.loc_name, psa_type, load_name,
                    load_desc, load_capacity, load_object_zone_id, 1)
        app.PrintPlain('Linking...: {}'.format(load_name))
        
        loads_data.append(row_data)
        
      
    # Clean table before insert
    connection.execute('DELETE FROM tb_lob_psaloads')
        
    connection.executemany( """insert into tb_lob_psaloads
    (fguid, filename, loadtype, loadid, loaddescription, 'loadcapacity',
    'object_fguid', 'cache_state') values (?,?,?,?,?,?,?,?)""", loads_data)
           
    connection.commit()
    

#Get user inputs
input_data = dict()
GUI()

mv_shapefile = input_data['MV Stations']
gis_shapefile = input_data['Load Zones']
dbfile = input_data['GLF File']
method = input_data['Method']
username = input_data['User']
password = input_data['Password']
base_year = input_data['Base Year']

#Establish SQLite connection to GLF file
connection = sqlite3.connect(dbfile)

# Set Base Year
#connection.execute('UPDATE tb_sys_properties SET baseyear = {}'.format(base_year))
#connection.commit()

# Constatnts
load_zones_table = 'tb_lob_loadobject'
psa_table = 'tb_lob_psaloads'
psa_type ='PowerFactory'

# PowerFactory actions
app = powerfactory.GetApplication()
project = app.GetActiveProject()

app.PrintPlain('Collecting loads from Project: {}'.format(project))
loads = app.GetCalcRelevantObjects('*.ElmLod')

if method == 'Use SharePoint data':
    app.PrintPlain('Downloading data')
    object_zone_ids = get_shapepoint_data(username, password)
else:
    app.PrintPlain('Analysing shapefiles...Please wait')
    object_zone_ids = get_linked_zones(gis_shapefile, mv_shapefile)

app.PrintPlain('Linking Loads')
link_loads(loads)

#Close db connection
connection.close()

app.PrintPlain('Done. Deactiving Project: {}'.format(project))
project.Deactivate()
