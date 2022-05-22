from flask import Flask

import numpy as np

import pandas as pd

from openpyxl import load_workbook

import copy

import os

from flask import send_from_directory, send_file, current_app, render_template

from collections import defaultdict

from dash import Dash, dcc, html, dash_table
import dash_bootstrap_components as dbc
from dash.dash_table.Format import Format, Padding

# Loading Workbook and assigning 'Sheet' as active
workbook = load_workbook('TestFile_051522.xlsx')
sheet = workbook.active

# Assigning cells into the sheets as variable xxxvalues
runvalues = sheet['A4:A23']
HRvalues = sheet['A26:A45']
RBIvalues = sheet['A48:A67']
SBvalues = sheet['A70:A89']
AVGvalues = sheet['A92:A111']
OPSvalues = sheet['A114:A133']
Wvalues = sheet['A137:A156']
QSvalues = sheet['A159:A178']
SVHvalues = sheet['A181:A200']
Kvalues = sheet['A203:A222']
ERAvalues = sheet['A225:A244']
WHIPvalues = sheet['A247:A266']

# Creating empty lists to be appended to in following function
runvalueslist = []
HRvalueslist = []
RBIvalueslist = []
SBvalueslist = []
AVGvalueslist = []
OPSvalueslist = []
Wvalueslist = []
QSvalueslist = []
SVHvalueslist = []
Kvalueslist = []
ERAvalueslist = []
WHIPvalueslist = []

# RUNS
# Captures all values (team names and stat values) for Runs category
for row in runvalues:
    for cell in row:
        runcellvalue = str(cell.value)
        runvalueslist.append(runcellvalue)

# Creates empty lists, then a for loop to append a cell to the list if it is even or odd (even = stat, odd = teamname)
runstatslist = []
runteamslist = []
for index in range(0, len(runvalueslist)):
    if index % 2 == 0:
        runstatslist.append(runvalueslist[index])
    else:
        runteamslist.append(runvalueslist[index])

# Next we'll need to combine them into a dictionary
runs_dict = {}
for runkey in runteamslist:
    for runvalue in runstatslist:
        runs_dict[runkey] = runvalue
        runstatslist.remove(runvalue)
        break

# HR
for row in HRvalues:
    for cell in row:
        HRcellvalue = str(cell.value)
        HRvalueslist.append(HRcellvalue)

HRstatslist = []
HRteamslist = []
for index in range(0, len(HRvalueslist)):
    if index % 2 == 0:
        HRstatslist.append(HRvalueslist[index])
    else:
        HRteamslist.append(HRvalueslist[index])

HRs_dict = {}
for HRkey in HRteamslist:
    for HRvalue in HRstatslist:
        HRs_dict[HRkey] = HRvalue
        HRstatslist.remove(HRvalue)
        break

# RBI
for row in RBIvalues:
    for cell in row:
        RBIcellvalue = str(cell.value)
        RBIvalueslist.append(RBIcellvalue)

RBIstatslist = []
RBIteamslist = []
for index in range(0, len(RBIvalueslist)):
    if index % 2 == 0:
        RBIstatslist.append(RBIvalueslist[index])
    else:
        RBIteamslist.append(RBIvalueslist[index])

RBIs_dict = {}
for RBIkey in RBIteamslist:
    for RBIvalue in RBIstatslist:
        RBIs_dict[RBIkey] = RBIvalue
        RBIstatslist.remove(RBIvalue)
        break

# SB
for row in SBvalues:
    for cell in row:
        SBcellvalue = str(cell.value)
        SBvalueslist.append(SBcellvalue)

SBstatslist = []
SBteamslist = []
for index in range(0, len(SBvalueslist)):
    if index % 2 == 0:
        SBstatslist.append(SBvalueslist[index])
    else:
        SBteamslist.append(SBvalueslist[index])

SBs_dict = {}
for SBkey in SBteamslist:
    for SBvalue in SBstatslist:
        SBs_dict[SBkey] = SBvalue
        SBstatslist.remove(SBvalue)
        break

# AVG
for row in AVGvalues:
    for cell in row:
        AVGcellvalue = cell.value
        AVGvalueslist.append(AVGcellvalue)

AVGstatslist = []
AVGteamslist = []
for index in range(0, len(AVGvalueslist)):
    if index % 2 == 0:
        AVGstatslist.append(AVGvalueslist[index])
    else:
        AVGteamslist.append(AVGvalueslist[index])

AVGs_dict = {}
for AVGkey in AVGteamslist:
    for AVGvalue in AVGstatslist:
        AVGs_dict[AVGkey] = AVGvalue
        AVGstatslist.remove(AVGvalue)
        break

# OPS
for row in OPSvalues:
    for cell in row:
        OPScellvalue = str(cell.value)
        OPSvalueslist.append(OPScellvalue)

OPSstatslist = []
OPSteamslist = []
for index in range(0, len(OPSvalueslist)):
    if index % 2 == 0:
        OPSstatslist.append(OPSvalueslist[index])
    else:
        OPSteamslist.append(OPSvalueslist[index])

OPSs_dict = {}
for OPSkey in OPSteamslist:
    for OPSvalue in OPSstatslist:
        OPSs_dict[OPSkey] = OPSvalue
        OPSstatslist.remove(OPSvalue)
        break

# W
for row in Wvalues:
    for cell in row:
        Wcellvalue = str(cell.value)
        Wvalueslist.append(Wcellvalue)

Wstatslist = []
Wteamslist = []
for index in range(0, len(Wvalueslist)):
    if index % 2 == 0:
        Wstatslist.append(Wvalueslist[index])
    else:
        Wteamslist.append(Wvalueslist[index])

Ws_dict = {}
for Wkey in Wteamslist:
    for Wvalue in Wstatslist:
        Ws_dict[Wkey] = Wvalue
        Wstatslist.remove(Wvalue)
        break

# QS
for row in QSvalues:
    for cell in row:
        QScellvalue = str(cell.value)
        QSvalueslist.append(QScellvalue)

QSstatslist = []
QSteamslist = []
for index in range(0, len(QSvalueslist)):
    if index % 2 == 0:
        QSstatslist.append(QSvalueslist[index])
    else:
        QSteamslist.append(QSvalueslist[index])

QSs_dict = {}
for QSkey in QSteamslist:
    for QSvalue in QSstatslist:
        QSs_dict[QSkey] = QSvalue
        QSstatslist.remove(QSvalue)
        break

# SVH
for row in SVHvalues:
    for cell in row:
        SVHcellvalue = str(cell.value)
        SVHvalueslist.append(SVHcellvalue)

SVHstatslist = []
SVHteamslist = []
for index in range(0, len(SVHvalueslist)):
    if index % 2 == 0:
        SVHstatslist.append(SVHvalueslist[index])
    else:
        SVHteamslist.append(SVHvalueslist[index])

SVHs_dict = {}
for SVHkey in SVHteamslist:
    for SVHvalue in SVHstatslist:
        SVHs_dict[SVHkey] = SVHvalue
        SVHstatslist.remove(SVHvalue)
        break

# K
for row in Kvalues:
    for cell in row:
        Kcellvalue = str(cell.value)
        Kvalueslist.append(Kcellvalue)

Kstatslist = []
Kteamslist = []
for index in range(0, len(Kvalueslist)):
    if index % 2 == 0:
        Kstatslist.append(Kvalueslist[index])
    else:
        Kteamslist.append(Kvalueslist[index])

Ks_dict = {}
for Kkey in Kteamslist:
    for Kvalue in Kstatslist:
        Ks_dict[Kkey] = Kvalue
        Kstatslist.remove(Kvalue)
        break

# ERA
for row in ERAvalues:
    for cell in row:
        ERAcellvalue = str(cell.value)
        ERAvalueslist.append(ERAcellvalue)

ERAstatslist = []
ERAteamslist = []
for index in range(0, len(ERAvalueslist)):
    if index % 2 == 0:
        ERAstatslist.append(ERAvalueslist[index])
    else:
        ERAteamslist.append(ERAvalueslist[index])

ERAs_dict = {}
for ERAkey in ERAteamslist:
    for ERAvalue in ERAstatslist:
        ERAs_dict[ERAkey] = ERAvalue
        ERAstatslist.remove(ERAvalue)
        break

# WHIP
for row in WHIPvalues:
    for cell in row:
        WHIPcellvalue = str(cell.value)
        WHIPvalueslist.append(WHIPcellvalue)

WHIPstatslist = []
WHIPteamslist = []
for index in range(0, len(WHIPvalueslist)):
    if index % 2 == 0:
        WHIPstatslist.append(WHIPvalueslist[index])
    else:
        WHIPteamslist.append(WHIPvalueslist[index])

WHIPs_dict = {}
for WHIPkey in WHIPteamslist:
    for WHIPvalue in WHIPstatslist:
        WHIPs_dict[WHIPkey] = WHIPvalue
        WHIPstatslist.remove(WHIPvalue)
        break

###################

# Creation of a 'defaultdict' that combines dictionaries on one key
statsdd = defaultdict(list)

for d in (runs_dict, HRs_dict, RBIs_dict, SBs_dict, AVGs_dict, OPSs_dict, Ws_dict, QSs_dict, SVHs_dict, Ks_dict, ERAs_dict, WHIPs_dict):
    for key, value in d.items():
        statsdd[key].append(value)

# Turns the defaultdict to 2 lists
statsdf_list_keys = list(statsdd.keys())
statsdf_list_values = list(statsdd.values())

# Combining the two lists above
statsdf_list = np.column_stack((statsdf_list_keys, statsdf_list_values))

# ## Creation of DataFrame, that formats it into a nice table, then assigning column headers
statsdf = pd.DataFrame(statsdf_list)
statsdf.columns = ['Team', 'R', 'HR', 'RBI', 'SB', 'AVG', 'OPS', 'W', 'QS', 'SV+H', 'K', 'ERA', 'WHIP']

# Using Dash to create a DataTable out of statsdf
app = Dash(__name__)
server = app.server

app.layout = html.Div(
    [
        html.Link(
            rel='stylesheet',
            href='/static/stylesheet.css'
        ),

        html.Div(
            className='title',
            children='The Show - Season Stats'
        ),
        html.Div(
            dash_table.DataTable(
                statsdf.to_dict('records'),
                id='table',
                columns=[{'id': c, 'name': c, 'type': 'text'} for c in statsdf.columns],
                sort_action='native',
                style_data_conditional=[
                    {
                        'if': {
                            'column_id': 'Team'
                        },
                        'textAlign': 'left',
                        'width': '10%',
                        'paddingLeft': '1%'

                    }
                ],
                style_data={'fontFamily': 'Nunito Sans', 'width': '5%', 'textAlign': 'center', 'height': '40px'},
                style_header_conditional=[
                    {
                        'if': {
                            'column_id': 'Team'
                        },
                        'textAlign': 'left',
                        'paddingLeft': '1%'
                    }
                ],
                style_header={'fontFamily': 'Nunito Sans', 'fontWeight': 'bold', 'textAlign': 'center', 'height': '10%'}
            )
        )
    ]
)

# Running the application on web server
if __name__ == '__main__':
    app.run_server(debug=True)