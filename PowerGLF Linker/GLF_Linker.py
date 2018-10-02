# Script to link PowerFactory MV casefile loads to PowerGLF
# 02/10/2018 1.3 Peace Lekalakala
import datetime
import os
import powerfactory
import sqlite3
import sys
import uuid

from functools import partial
from fuzzywuzzy import process, fuzz
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

url = 'https://mysites.eskom.co.za/personal/elec_lekalap/Planning/linkedzones/lz_table.aspx'

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
                    
    def open_link(event=None):
        import webbrowser
        webbrowser.open_new(url)
        
    def set_values(method):
        input_data['MV Stations'] = e2.get()
        input_data['Load Zones'] = e3.get()
        input_data['GLF File'] = e1.get()
        input_data['User'] = e4.get()
        input_data['Password'] = e5.get()
        input_data['Method'] = method
        #input_data['Base Year'] = e6.get()
        
        # Exit gui
        main.destroy()
        
    def cancel_op():
        main.destroy()
        sys.exit()
                    
                
    main = Tk()
    main.title('PowerGLF Linker')
    main.geometry('900x500')
    
    glf_gis_shapefile = StringVar()
    mv_shapefile = StringVar()
    glf_file = StringVar()
    user_name = StringVar()
    user_name.set(os.getenv('username'))
    
    Label(main, font="Bold", text="Select Checked out GLF file:").place(x=20, y=65)
    e1 = Entry(main, width=95, textvariable=glf_file)
    e1.place(x=20, y=95)
    b1 = Button(main, text="Browse...", command=get_glf_file)
    b1.place(x=805, y=92)
    
    # gives weight to the cells in the grid
    rows = 0
    while rows < 50:
        main.rowconfigure(rows, weight=1)
        main.columnconfigure(rows, weight=1)
        rows += 1
 
    # Defines and places the notebook widget
    nb = ttk.Notebook(main)
    nb.grid(row=18, column=1, columnspan=48, rowspan=30, sticky='NESW')
 
    # Adds tab 1 of the notebook
    page1 = ttk.Frame(nb)
    nb.add(page1, text='Use Shapefiles')
 
    # Adds tab 2 of the notebook
    page2 = ttk.Frame(nb)
    nb.add(page2, text='Use pre compiled data')

    Label(page1, font="Bold", text="MV Stations Shapefile:").place(x=10, y=25)
    e2 = Entry(page1, width=95, textvariable=mv_shapefile)
    e2.place(x=10, y=60)
    b2 = Button(page1, text="Browse...", command=get_mv_shape)
    b2.place(x=795, y=55)
    
    Label(page1, font="Bold", text="GLF GIS Land Use Shapefile:").place(x=10, y=100)
    e3 = Entry(page1, width=95, textvariable=glf_gis_shapefile)
    e3.place(x=10, y=135)
    b3 = Button(page1, text="Browse...", command=get_landuse_shape)
    b3.place(x=795, y=130)

    Button(page1, font="Bold", text="Cancel", width=35, command=cancel_op).place(x=20, y=225)
    page1_action = partial(set_values, 'Use imported shapefiles')
    Button(page1, font="Bold", text="Link Loads", width=35, command=page1_action).place(x=455, y=225)
    
    Label(page2, text="Enter your sharepoint login credentials below").place(x=10, y=20)
    Label(page2, font="Bold", text="UserName:").place(x=10, y=50)
    e4 = Entry(page2, width=50, textvariable=user_name)
    e4.place(x=10, y=85)
    
    Label(page2, font="Bold", text="Password:").place(x=450, y=50)
    e5 = Entry(page2, show="*", width=50)
    e5.place(x=450, y=85)
            
    link = Label(page2, font="Bold", text="View SharePoint data",
                 underline=True, fg="blue", cursor="hand2")
    link.place(x=10, y=140)
    link.bind("<Button-1>", open_link)
    
    Button(page2, font="Bold", text="Cancel", width=35,
           command=cancel_op).place(x=20, y=225)
    page2_action = partial(set_values, 'Use SharePoint data')
    Button(page2, font="Bold", text="Link Loads", width=35,
           command=page2_action).place(x=455, y=225)
 
    main.mainloop()
    

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

    
def create_geoframe(shape_file):
    import shapefile
    from shapely.geometry import Point
    import pandas as pd
    import geopandas as gpd
    
    sf = shapefile.Reader(shape_file)
    shapes = sf.shapes()
    shape_records = sf.shapeRecords()
    
    fields = sf.fields[1:]
    field_names = [field[0] for field in fields]
    records = [i.record for i in shape_records]
    points = [i.points for i in shapes]
    #print(points)
    df = pd.DataFrame(records, columns=field_names)
    
    geometry = [Point(i[0]) for i in points]
    crs = {'init': 'epsg:4326'}
    
    return gpd.GeoDataFrame(df, geometry=geometry, crs=crs)
    
    
def get_linked_zones(gis_shapefile, mv_shapefile, study_feeders):
    import geopandas as gpd
    
    glf_df = gpd.read_file(gis_shapefile)
    glf_df = glf_df.to_crs({'init': 'epsg:4326'})
    
    try:
        mv_df = gpd.read_file(mv_shapefile)
        mv_df = mv_df.to_crs({'init':'epsg:4326'})
    except ValueError:
        mv_df = create_geoframe(mv_shapefile)

    load_zones_names = connection.execute('select id from {}'.format(load_zones_table))
    load_zones_names = [i[0] for i in load_zones_names.fetchall()]

    zone_ids = get_object_ids(load_zones_table, 'fguid')
    load_zones_glf = dict(zip(load_zones_names, zone_ids))

    # Get corresponding load/transformers zone
    merged_df = gpd.sjoin(mv_df, glf_df, how="left", op="within")
    merged_df['PF Name'] = merged_df.standard_l.apply(lambda x: get_pf_name(x))
    
    feeder_names = [i for i in list(merged_df['parent_des'].unique())
                    if 'Feeder' in i]
    
    feeders_ = [process.extractOne(i, feeder_names)[0] for i in study_feeders]
    merged_df = merged_df[merged_df['parent_des'].isin(feeders_)]
    identified_zones = dict(zip(merged_df['PF Name'], merged_df.lob_id))
    
    object_zone_ids = {}
    for k, v in identified_zones.items():
        try:
            object_zone_ids[k] = load_zones_glf[v]
        except KeyError:
            pass
    
    return object_zone_ids
    

def get_sharepoint_data(username, password, study_feeders):
    from requests_ntlm import HttpNtlmAuth
    import requests
    import pandas as pd
    import lxml.html as LH
        
    auth = HttpNtlmAuth('elec\\' + username, password)
    html = requests.get(url, auth=auth)
    app.PrintPlain(html)
    html_data = html.text
    
    table = LH.fromstring(html_data)
    # extract the text from `<td>` tags
    data = [[elt.text_content() for elt in tr.xpath('td')] 
            for tr in table.xpath('//tr')]
    headers = ['Sector', 'CNC', 'Feeder', 'Plant Slot ID', 'Description',
               'Label', 'Load Zone', 'Subclass', 'Zone ID']
    df = pd.DataFrame(data, columns=headers)
    df = df.reindex(df.index.drop(0))
    app.PrintPlain(df)
    df['PF Name'] = df['Label'].apply(lambda x: get_pf_name(x))
    feeder_names = [i for i in list(df.Feeder.unique()) if 'Feeder' in i]
    
    feeders_ = [process.extractOne(i, feeder_names)[0] for i in study_feeders]
    df = df[df['Feeder'].isin(feeders_)]
    object_zone_ids = dict(zip(df['PF Name'], df['Zone ID']))
    
                
    return object_zone_ids
    
    
   
def collect_feeder_zones():
    feeder_ids = get_object_ids('tb_lob_node', 'fguid')
    nodes = connection.execute('select id from tb_lob_node').fetchall()
    nodes = [i[0] for i in nodes]
    
    feeder_zones = {}
    for i, j in enumerate(nodes):
        if 'Feeder' in j:
            feeder_zones[j] = feeder_ids[i]
    
    return feeder_zones
    
           
    
def get_load_zone(load):
    app.PrintInfo('Processing: {}'.format(load))
    load_short_name = load.fold_id.loc_name
    if load_short_name in object_zone_ids.keys():
        zone_id = object_zone_ids[load_short_name]
    else:
        try:
            load_sw_name = load.desc[0]
            feeder, mv_station = str(load_sw_name).split(' Feeder ')[:2]
            _station = mv_station.split()[0].replace('/', ' ').replace('-', ' ')
            zone_id = object_zone_ids[_station]
        except IndexError:
            load_sw_name = load.fold_id.desc[0]
            possible_match = process.extractOne(load_sw_name, zone_names)
            if possible_match[1] >= 70:
                _station = possible_match[0]
                zone_id = object_zone_ids[_station]
            else:
                zone_id = ''
        except KeyError:
            possible_match = process.extractOne(load_short_name, zone_names)
            if possible_match[1] >= 70:
                _station = possible_match[0]
                zone_id = object_zone_ids[_station]
            else:
                zone_id = ''
            
        else:
            zone_id = ''
    
    return zone_id
    
    
def get_feeder_zone(load):
    # As per guideline, assign loads to feeder zones
    name = load.fold_id.loc_name
    try:
        load_sw_name = load.desc[0]
    except IndexError:
        load_sw_name = load.fold_id.desc[0]
    
    feeder_ = str(load_sw_name).split(' Feeder ')[0]
    
    try:
        feeder = loads_supplying_feeders[feeder_]
        id = feeder_zones[feeder]
    except KeyError:
        feeder = process.extractOne(feeder_, feeder_zones.keys())[0]
        loads_supplying_feeders[feeder_] = feeder
        id = feeder_zones[feeder]
    app.PrintInfo('assigning load: {0} to {1} feeder node'.format(load.loc_name, feeder))
    return id
    
   
    
def link_loads(loads):
    # Link to feeder first
    loads_data = [('load', project.loc_name, psa_type, i.GetFullName(),
                   i.loc_name, 0, get_feeder_zone(i), 1) for i in loads]
       
    # Clean table before insert
    connection.execute('DELETE FROM tb_lob_psaloads')
    connection.commit()
    
    # Create UniqueIDs for fguid with a trigger
    try:
        connection.execute("DROP TRIGGER new_uid")
    except Exception:
        pass
            
    connection.execute('''CREATE TRIGGER new_uid AFTER INSERT ON tb_lob_psaloads
        FOR EACH ROW
        WHEN (NEW.fguid IS NOT NULL)
        BEGIN
            UPDATE tb_lob_psaloads SET fguid = (select hex( randomblob(4)) || '-' || hex( randomblob(2))
             || '-' || '4' || substr( hex( randomblob(2)), 2) || '-'
             || substr('AB89', 1 + (abs(random()) % 4) , 1)  ||
             substr(hex(randomblob(2)), 2) || '-' || hex(randomblob(6)) ) WHERE rowid = NEW.rowid;
         END;''')
         
    
    # Insert and assign loads to a load object on the database
    connection.executemany( """insert into tb_lob_psaloads
    (fguid, filename, loadtype, loadid, loaddescription, 'loadcapacity',
    'object_fguid', 'cache_state') values (?,?,?,?,?,?,?,?)""", loads_data)
    connection.commit()
    
    
def assign_to_loadzones():
    
    app.PrintInfo('Getting load zones...')
    
    # Link loads to respective load zones
    load_zones_data = []
    for load in loads:
        load_zone_id = get_load_zone(load)
        if load_zone_id:
            row_data = (load.GetFullName(), load_zone_id)
            load_zones_data.append(row_data)
        else:
            pass
                
    # Update loads with their respective load zones     
    connection.executemany("""UPDATE tb_lob_psaloads 
        SET object_fguid=? 
        WHERE loadid=?""", load_zones_data)
    connection.commit()  
    
    # Remove trigger and save
    connection.execute("DROP TRIGGER new_uid")
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


#Establish SQLite connection to GLF file
connection = sqlite3.connect(dbfile)

# Constatnts
load_zones_table = 'tb_lob_loadobject'
psa_table = 'tb_lob_psaloads'
psa_type ='PowerFactory'

# PowerFactory actions
app = powerfactory.GetApplication()
project = app.GetActiveProject()

app.PrintInfo('Collecting loads from Project: {}'.format(project))
loads = app.GetCalcRelevantObjects('*.ElmLod')

#Identify all feeder nodes/zones in your PowerGLF file
feeder_zones = collect_feeder_zones()
loads_supplying_feeders = {}

# Begin linking all loads to their respective feeder nodes in PowerGLF
app.PrintPlain('Linking Loads')
link_loads(loads)


study_feeders = loads_supplying_feeders.values()
if method == 'Use SharePoint data':
    app.PrintInfo('Downloading load zones data...Please wait')
    object_zone_ids = get_sharepoint_data(username, password, study_feeders)
else:
    app.PrintPlain('Analysing shapefiles...Please wait')
    object_zone_ids = get_linked_zones(gis_shapefile, mv_shapefile, study_feeders)
    
zone_names = list(object_zone_ids.keys())
assign_to_loadzones()

#Close db connection
connection.close()

app.PrintPlain('Done. Deactiving Project: {}'.format(project))
project.Deactivate()
