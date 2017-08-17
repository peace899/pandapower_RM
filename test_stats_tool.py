# MV90 Stats Tool
# SElect substation
# select feeder/meter
# Select period

import dash
import dash_core_components as dcc
import dash_html_components as html

import pandas as pd
import flask
import os

from datetime import datetime

"""
open('//HOST/share/path/to/file')
open(r'\\HOST\share\path\to\file')
or
open('\\\\HOST\\share\\path\\to\\file')
#\\ecrfnp01\USERG01\Sharedat\Planning_Central\Planning_Users\Peace

"""
csv_file = r'C:\Users\lekalap\Documents\Personal\scripts\JoburgVaal.csv'
#Stats data
#csv_file = r'C:\Users\lekalap\Desktop\stats\allstats.csv'
#csv_file = r'C:\Users\Peace\Documents\scripts\eskom\MIdrand\CROWTHORNE\Witpoortjie.csv'
df = pd.read_csv(csv_file)
# filter results by recorder id
#df1 = df[df['RECORDER ID'] == 'JBS-1171122760']
#df1 = '' 
# filter results by date

#df2 = df1[df['DATE'] == '2016/01/01']
#y = df2['kVA'].to_string(index=False).split('\n')
#x = [i.split()[1] for i in df2['DATE_TIME'].to_string(index=False).split('\n')]
# Substations_feeders list
subs_file = 'subslist.csv'
df_subs = pd.read_csv(subs_file)
subs = df_subs['Substation'].tolist()
sub_fdr_combos = {}
for sub in sorted(list(set(subs))):
    df_sub = df_subs[df_subs['Substation'] == sub]
    feeders = df_sub['Location'].to_string(index=False).split('\n')
    recorders = df_sub['RecorderID'].to_string(index=False).split('\n')
    combo = [(i,j) for i,j in zip(feeders, recorders)]
    sub_fdr_combos[sub] = combo


def stats_data(recoder_id):
    # filter results by recorder id
    df1 = df[df['RECORDER ID'] == recoder_id]
    return df1

def get_months(dataframe):
    dates = dataframe['DATE'].unique()
    m_y_choices = []
    for date in dates:
        m_y_choice = datetime.strptime(date, "%Y/%m/%d").strftime('%b %Y')
        if m_y_choice not in m_y_choices:
            m_y_choices.append(m_y_choice)
    return m_y_choices
    
def monthly_highs_stats(dataframe):
    months = []
    for date in dataframe['DATE'].unique():
        month_year = datetime.strptime(date, "%Y/%m/%d").strftime('%b %Y')
        if month_year not in months:
            months.append(month_year)
    maximums = []
    for month in months:
        m, y = month.split()
        monthly_max = (month, 
                       dataframe[(dataframe['MONTH'] == m) &
                       (dataframe['DATE'].str.contains(y))]['kVA'].max())
        maximums.append(monthly_max)
    months = [i[0].split()[0] for i in maximums]
    loads = [i[1] for i in maximums]
    new_df = dataframe[(dataframe['MONTH'].isin(months) &
             dataframe['kVA'].isin(loads))]
    return maximums, new_df       

def daily_highs_stats(dataframe):
    maximums = []
    days = dataframe['DATE'].unique()
    for day in days:
        daily_max = day, dataframe[dataframe['DATE'] == day]['kVA'].max()
        maximums.append(daily_max)
    idx = dataframe.groupby(['DATE'])['kVA'].transform(max) == dataframe['kVA']
    new_df = dataframe[idx]
    return maximums, new_df
    
def single_day_stats(dataframe, specific_date):
    hourly_loads = []
    df_day = dataframe[dataframe['DATE'] == specific_date]
    day_times = df_day['DATE_TIME'].to_string(index=False).split('\n')
    for day_time in day_times:
        hour_load = (day_time.split()[1],
                     dataframe.loc[dataframe['DATE_TIME'] == day_time]['kVA'].
                     values[0])
        hourly_loads.append(hour_load)
    return hourly_loads, df_day

def single_month_stats(dataframe, specific_month):
    m, y = specific_month.split()
    df_month = dataframe[(dataframe['MONTH'] == m) &
               (dataframe['DATE'].str.contains(y))]
    
def generate_table(dataframe):
    return html.Table([
                html.Thead([
                    html.Tr([
                       html.Th(col) for col in dataframe.columns],
                       className="cell")],
                       className="row header",
                       style={'font-size': '95%'}),
                html.Tbody([
                    html.Tr([
                        html.Td(
                        dataframe.iloc[i][col]) for col in dataframe.columns],
                        className="cell",
                        style={
                        'font-size': '85%'}) for i in range(len(dataframe))],
                        className="row")    
            ],className="table table--fixed")

def generate_graph(dframe):
    y = dframe['kVA'].to_string(index=False).split('\n')
    #x = [i.split()[0] for i in dataframe['DATE_TIME'].to_string(index=False).split('\n')]
    x = dframe['DATE_TIME'].to_string(index=False).split('\n')
    return html.Div([dcc.Graph(id='my-graph',
                               figure={'data': [{'y': y, 'x': x}]})],
                               style={'width': '800', 'height': '360px',
                                      'float': 'right',
                                      'margin-top': '-425px'})

def graphic_show(dataframe):
    table = generate_table(dataframe)
    figure = generate_graph(dataframe)
    return html.Div([table, figure])
    
app = dash.Dash()

Output = dash.dependencies.Output  
Input = dash.dependencies.Input

app.layout = html.Div([

    html.H1('MV90 Stats Data Tool'),
    
    html.Hr(),
    
    html.Div([
        html.Label('Select Substation:'),
        dcc.Dropdown(id='station_dropdown',
            options=[{'label': k,
                      'value': k} for k,v in sorted(sub_fdr_combos.items())])]
        ,style={'width': '25%', 'float': 'left', 'display': 'inline-block'}),
    html.Div([
        html.Label('Select Feeder/Meter Point:'),
        dcc.Dropdown(id='feeder_dropdown')]                      
        ,style={'width': '30%', 'position': 'relative', 'left': '60px',
                 'display': 'inline-block'}),
    
    html.Div([
        html.Label('RecorderID:')],
        style={'position': 'relative',
               'left': '80px', 'display': 'inline-block'}),
    
    html.Div([
        dcc.Input(id='display-selected-values')]                      
        ,style={'height': '15px', 'position': 'relative', 'left': '0px',
                 'top': '30px', 'display': 'inline-block'}),
    html.Hr(),

    
    html.Div(id='my-graphic'),
             
      
    
    
    html.Hr()
    
    
    ])

@app.callback(
    dash.dependencies.Output('feeder_dropdown', 'options'),
    [dash.dependencies.Input('station_dropdown', 'value')])
def set_feeder_options(selected_sub):
    return [{'label': i[0],
             'value': i[0]} for i in sub_fdr_combos[selected_sub]]

@app.callback(
    dash.dependencies.Output('feeder_dropdown', 'value'),
    [dash.dependencies.Input('feeder_dropdown', 'options')])
def set_feeders_value(available_options):
    return available_options[0]['value']

@app.callback(
    dash.dependencies.Output('display-selected-values', 'value'),
    [dash.dependencies.Input('station_dropdown', 'value'),
     dash.dependencies.Input('feeder_dropdown', 'value')])
def set_display_children(selected_sub, selected_feeder):
    for x in sub_fdr_combos[selected_sub]:
        if selected_feeder in x:
            selected_meter = x[1]
            return selected_meter


@app.callback(
    dash.dependencies.Output('my-graphic', 'children'),
    [dash.dependencies.Input('display-selected-values', 'value')])
def set_initial_table(rec_id):
    df1 = stats_data(rec_id)
    return [graphic_show(df1)]
    #graphic_show(stats_data(rec_id))
    #return df1
"""     
@app.callback(
    dash.dependencies.Output('my-table', 'children'),
    [dash.dependencies.Input(df1, 'value')])
def set_initial_table(rec_id):
    df1 = stats_data(rec_id)
    table = generate_table(df1)    
    return table      

@app.callback(
    dash.dependencies.Output('my-graph', 'children'),
    [dash.dependencies.Input(df1, 'value')])
def set_initial_graph(hrec_id):
    df1 = stats_data(rec_id)
    figure = generate_graph(df1)
    return figure

@app.callback(Output('intermediate-value', 'children'), [Input('dropdown', 'value')])
def stats_data(value):
     # some expensive clean data step
     cleaned_df = your_expensive_clean_or_compute_step(value)
     return cleaned_df.to_json() # or, more generally, json.dumps(cleaned_df)

@app.callback(Output('graph', 'figure'), [Input('intermediate-value', 'children'])
def update_graph(jsonified_cleaned_data):
    dff = pd.read_json(jsonified_cleaned_data) # or, more generally json.loads(jsonified_cleaned_data)
    figure = create_figure(dff) 
    return figure

@app.callback(Output('table', 'children'), [Input('intermediate-value', 'children'])
def update_table(jsonified_cleaned_data):
    dff = pd.read_json(jsonified_cleaned_data) # or, more generally json.loads(jsonified_cleaned_data)
    table = create_table(dff) 
    return table
"""
css_directory = os.getcwd()
stylesheets = ['stylesheet.css']
static_css_route = '/static/'


@app.server.route('{}<stylesheet>'.format(static_css_route))
def serve_stylesheet(stylesheet):
    if stylesheet not in stylesheets:
        raise Exception(
            '"{}" is excluded from the allowed static files'.format(
                stylesheet
            )
        )
    return flask.send_from_directory(css_directory, stylesheet)


for stylesheet in stylesheets:
    app.css.append_css({"external_url": "/static/{}".format(stylesheet)})

app.config.suppress_callback_exceptions=True
#app.css.append_css({"external_url": "https://codepen.io/chriddyp/pen/bWLwgP.css"})   
#app.css.config.serve_locally = True
#app.scripts.config.serve_locally = True
if __name__ == '__main__':
    app.run_server(debug=True)