# -*- coding: UTF-8 -*-
"""
This script is to analyze aging records
Created on Thu Feb 14 17:50:00 2023
@author ACES ANALYTICS TEAM

"""

"""
Index
## 1. Create start time of running script
## 2. Import modules packages
## 3. Define date, path, file names
## 4. Import raw data
## 5. Create string of date
## 6. Manipulate raw data
## 7. Plot
## 8. Format worksheets
## 9. Export results to Excel file
"""

"""
## 1. Create start time of running script
"""
# Create start time of script running
from timeit import default_timer as timer
start = timer()

"""
## 2. Import modules, packages
"""
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Qt5Agg")  # Do this before importing pyplot.
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from xlwings.utils import rgb_to_int

"""
## 3. Define date, path, file names
"""
# Define path
path_from = r'D:\Python\ACES_Analytics\ACES001\Output'
path_to = r'D:\Python\ACES_Analytics\ACES006\Output'

# Define file names
input_file = path_from + "\\" + "ACES001_Sales Records.xlsx"
input_file1 = path_from + "\\" + "ACES001_Sales Budget.xlsx"
output_file = path_to + "\\" + "ACES006_Variance Analysis Report.xlsx"

"""
## 4. Import raw data 
"""
raw_data_act = pd.read_excel(input_file)
raw_data_bud = pd.read_excel(input_file1)

"""
## 5. Create string of date
"""
from datetime import datetime
import calendar
#import datetime
raw_date = raw_data_act.iat[0,1]

str_mth_year = datetime.strptime(raw_date,"%Y-%m-%d")
last_date = str_mth_year.replace(day = calendar.monthrange(str_mth_year.year, str_mth_year.month)[1])
Last_date_str = datetime.strftime(last_date,"%d-%b-%y")
mth_str = datetime.strftime(last_date,"%b-%y")

"""
## 6. Manipulate raw data
"""
## Manipulate actual data
raw_data_act.at['Total', "Sales Revenue"] = raw_data_act["Sales Revenue"].sum()
raw_data_act.at['Total', "Sales Volume"] = raw_data_act["Sales Volume"].sum()
raw_data_act.at['Total', "Raw Material Cost"] = raw_data_act["Raw Material Cost"].sum()
raw_data_act.at['Total', "Fixed Exp"] = raw_data_act["Fixed Exp"].sum()
raw_data_act.at['Total', "Variable Exp"] = raw_data_act["Variable Exp"].sum()
raw_data_act.at['Total', "Operating Income"] = raw_data_act["Operating Income"].sum()
raw_data_act1 = raw_data_act.fillna("")

raw_data_act2 = raw_data_act1.loc[['Total'],:]

raw_data_act2 = raw_data_act2.loc[:,['Sales Revenue','Sales Volume',
       'Raw Material Cost', 'Fixed Exp', 'Variable Exp', 'Operating Income']]

raw_data_act2 = round(raw_data_act2 / 1000,2)

raw_data_act2['Unit Price'] = round(raw_data_act2['Sales Revenue'] / raw_data_act2['Sales Volume'],4)
raw_data_act2['Unit RMC'] = round(raw_data_act2['Raw Material Cost'] / raw_data_act2['Sales Volume'],4)

## Manipulate budget data
raw_data_bud.at['Total', "Sales Revenue"] = raw_data_bud["Sales Revenue"].sum()
raw_data_bud.at['Total', "Sales Volume"] = raw_data_bud["Sales Volume"].sum()
raw_data_bud.at['Total', "Raw Material Cost"] = raw_data_bud["Raw Material Cost"].sum()
raw_data_bud.at['Total', "Fixed Exp"] = raw_data_bud["Fixed Exp"].sum()
raw_data_bud.at['Total', "Variable Exp"] = raw_data_bud["Variable Exp"].sum()
raw_data_bud.at['Total', "Operating Income"] = raw_data_bud["Operating Income"].sum()
raw_data_bud1 = raw_data_bud.fillna("")

raw_data_bud2 = raw_data_bud1.loc[['Total'],:]

raw_data_bud2 = raw_data_bud2.loc[:,['Sales Revenue','Sales Volume',
       'Raw Material Cost', 'Fixed Exp', 'Variable Exp', 'Operating Income']]

raw_data_bud2 = round(raw_data_bud2 / 1000,2)

raw_data_bud2['Unit Price'] = round(raw_data_bud2['Sales Revenue'] / raw_data_bud2['Sales Volume'],4)
raw_data_bud2['Unit RMC'] = round(raw_data_bud2['Raw Material Cost'] / raw_data_bud2['Sales Volume'],4)
bud_operating_income = round(raw_data_bud2.at['Total','Operating Income'],0)
act_operating_income = round(raw_data_act2.at['Total','Operating Income'],0)

## Calculate variances
revenue_effect_due_to_price_change = round((raw_data_act2.at['Total','Unit Price'] - raw_data_bud2.at['Total','Unit Price']) \
                                     * raw_data_act2.at['Total','Sales Volume'],0)

revenue_effect_due_to_volume_change = round((raw_data_act2.at['Total','Sales Volume'] - raw_data_bud2.at['Total','Sales Volume']) \
                                     * raw_data_bud2.at['Total','Unit Price'],0)

rmc_effect_due_to_price_change = - round((raw_data_act2.at['Total','Unit RMC'] - raw_data_bud2.at['Total','Unit RMC']) \
                                     * raw_data_act2.at['Total','Sales Volume'],0)

rmc_effect_due_to_volume_change = - round((raw_data_act2.at['Total','Sales Volume'] - raw_data_bud2.at['Total','Sales Volume']) \
                                     * raw_data_bud2.at['Total','Unit RMC'],0)

fixed_exp_variance = - round(raw_data_act2.at['Total','Fixed Exp'] - raw_data_bud2.at['Total','Fixed Exp'],0)

variable_exp_variance = - round(raw_data_act2.at['Total','Variable Exp'] - raw_data_bud2.at['Total','Variable Exp'],0)

check = bud_operating_income + revenue_effect_due_to_price_change + revenue_effect_due_to_volume_change \
        + rmc_effect_due_to_price_change + rmc_effect_due_to_volume_change + fixed_exp_variance + variable_exp_variance \
        - act_operating_income

variable_exp_variance = - round(raw_data_act2.at['Total','Variable Exp'] - raw_data_bud2.at['Total','Variable Exp'],0) - check

print(check)
check1 = bud_operating_income + revenue_effect_due_to_price_change + revenue_effect_due_to_volume_change \
        + rmc_effect_due_to_price_change + rmc_effect_due_to_volume_change + fixed_exp_variance + variable_exp_variance \
        - act_operating_income

print(check1)

raw_data_act2 = raw_data_act2.rename(index = {'Total':'Act'})
raw_data_bud2 = raw_data_bud2.rename(index = {'Total':'Bud'})

variance_analysis_ot = pd.concat([raw_data_act2,raw_data_bud2])
variance_analysis_ot.loc['Act vs Bud'] = variance_analysis_ot.loc['Act'] - variance_analysis_ot.loc['Bud']

variance_amount = [bud_operating_income,revenue_effect_due_to_price_change,revenue_effect_due_to_volume_change,
                  rmc_effect_due_to_price_change, rmc_effect_due_to_volume_change,
                  fixed_exp_variance,variable_exp_variance,act_operating_income]

variance_factor = ["Bud OI","Revenue - Price Change", "Revenue - Volume Change",
                   "RMC - Price Change","RMC - Volume Change", "Fixed Exp Var",
                   "Var Exp Var","Act OI"]

variance_amount1 = [bud_operating_income,revenue_effect_due_to_price_change,revenue_effect_due_to_volume_change,
                  rmc_effect_due_to_price_change, rmc_effect_due_to_volume_change,
                  fixed_exp_variance,variable_exp_variance]

variance_factor1 = ["Bud OI","Revenue - Price Change", "Revenue - Volume Change",
                   "RMC - Price Change","RMC - Volume Change", "Fixed Exp Variance",
                   "Var Exp Variance"]

variance_factor2 = ["Bud Operating Income","Revenue Effect due to Price Change", "Revenue Effect due to Volume Change",
                   "RMC Effect due to Price Change","RMC Effect due to Volume Change", "Fixed Expense Variance",
                   "Variable Expense Variance","Actual Operating Income"]

variance_factors_amount = pd.DataFrame({"Factor":variance_factor2,
                                       "Amount":variance_amount})

variance_factors_amount_ot = variance_factors_amount.T

"""
## 7. Plot
"""
#Use python syntax to format currency
def amount(x, pos):
    'The two args are the value and tick position'
    return "k'${:,.0f}".format(x)
formatter = FuncFormatter(amount)

#Data to plot
index = variance_factor
raw_data = variance_amount
data = pd.Series(data = raw_data, index = index)

#Store data and create a blank series to use for the waterfall
#trans = pd.DataFrame(data=data,index=index)
trans = pd.DataFrame({'pos':np.maximum(data,0),'neg':np.minimum(data,0)},index = index) #,columns = 'amount')
#blank = trans.amount.cumsum().shift(1).fillna(0)
blank = data.cumsum().shift(1).fillna(0)

#The steps graphically show the levels as well as used for label placement
step = blank.reset_index(drop=True).repeat(3).shift(-1)
step[1::3] = np.nan

#When plotting the last element, we want to show the full bar,
#Set the blank to 0
blank.loc["Act OI"] = 0

#Plot and label

font_color = '#525252'

# Set blank background
plt.rcParams['axes.facecolor']='none'
plt.rcParams['savefig.facecolor']='none'

ax = trans.plot(kind='bar', stacked=True, bottom=blank, figsize=(10, 5),
                color = ['#70d1d4','#efd1f1']) #legend=None,  "#32beef" "#03aff4" '#efd1f1 pink'

ax.legend(loc = "center", bbox_to_anchor = (0,0.9,0.12,0),fontsize = 8)

ax.plot(step.index, step.values,'#52f253')#ffc700 #03aff4 #52f253

#ax.set_xlabel("Variance Factors", color = font_color, fontsize = 13,font = 'Verdana')
#ax.set_title(Title,color = font_color, fontsize = 13,font = 'Verdana')

#Format the axis for dollars
ax.yaxis.set_major_formatter(formatter)

#Rotate the labels
ax.set_xticklabels(trans.index,rotation=0)
plt.xticks(fontsize=8, color = font_color,font = 'Verdana')
plt.yticks(fontsize=12, color = font_color,font = 'Verdana')

## Get lables for the chart
#Data to plot. Do not include a total, it will be calculated
index = variance_factor1
data = {'amount': variance_amount1}
#data = {'amount': [800000,-80000,-75000,-85000,105000,-20000]}

#Store data and create a blank series to use for the waterfall
trans = pd.DataFrame(data=data,index=index)
blank = trans.amount.cumsum().shift(1).fillna(0)

#Get the net total number for the final element in the waterfall
total = trans.sum().amount
trans.loc["net"]= total
blank.loc["net"] = total

#Get the y-axis position for the labels
y_height = trans.amount.cumsum().shift(1).fillna(0)

#Get an offset so labels don't sit right on top of the bar
max = y_height.max()
neg_offset = max / 25 #25
pos_offset = max / 50 #50
plot_offset = int(max / 15) #15

#Start label loop
loop = 0
for index, row in trans.iterrows():
    # For the last item in the list, we don't want to double count
    if row['amount'] == total:
        y = y_height[loop]
    else:
        y = y_height[loop] + row['amount']
    # Determine if we want a neg or pos offset
    if row['amount'] > 0:
        y += pos_offset
    else:
        y -= neg_offset
    ax.annotate("{:,.0f}".format(row['amount']),(loop,y),ha="center",color = font_color,fontsize = 12, font="Verdana")
    loop+=1

#Scale up the y axis so there is room for the labels
ax.set_ylim(0,blank.max()+int(plot_offset))

# Remove axes splines
for s in ['top','bottom','left','right']:
    ax.spines[s].set_visible(False)

# Remove x,y Ticks
ax.xaxis.set_ticks_position('none')
ax.yaxis.set_ticks_position('none')

# Add padding between axes and labels
ax.xaxis.set_tick_params(pad=5.0)
ax.yaxis.set_tick_params(pad=5.0)

# Add text watermark

fig = ax.get_figure()
fig.text(0.5, 0.5, 'ACES Analytics Team', color = '#efd1f1',fontsize = 18,  ha = 'center', va = 'top',alpha = 0.5) #color = '#2ea7db'

# Change
width = 16
height = 6
fig.set_size_inches(width, height)

plt.show()

"""
## 8. Format worksheets with visualized plots
"""
# Open blank work book
import xlwings as xw
wb = xw.Book()

# region Format Budget Data
# Define name of worksheet
sheet = wb.sheets["Sheet1"]
sheet.name = "Bud"

# Hide gridlines for sheet
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Assign values of inventory records to worksheet
sheet.range("A1").options(index=False).value = raw_data_bud
sheet.autofit(axis="columns")

# Get length of columns, rows
row_len = len(raw_data_bud) + 1
row_len_minus1 = row_len - 1

clmn_len = len(raw_data_bud.columns)

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

name_last_clmn = excel_column_name(clmn_len)

# Format common features
first_cell = 'A1'
last_cell = name_last_clmn + str(row_len)

rng = sheet.range(first_cell, last_cell)
rng.row_height = 18
rng.api.Font.Name = 'Verdana'
rng.api.Font.Size = 10
rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
rng.api.WrapText = False

# Format data
first_cell = 'I1'
last_cell = name_last_clmn + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

first_cell = 'I1'
last_cell = "I" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'J2'
last_cell = "K" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'L2'
last_cell = "M" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'N2'
last_cell = "N" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'O2'
last_cell = "O" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'P2'
last_cell = "P" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'Q2'
last_cell = "R" + str(row_len)
rng = sheet.range(first_cell, last_cell)
rng.number_format = "#,###"

# Format first row
first_cell = 'A1'
last_cell = name_last_clmn + "1"
rng = sheet.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.WrapText = True
rng.color = ("#70d1d4")
rng.api.Font.Color = 0xffffff
rng.api.Font.Bold = True
rng.row_height = 20

# Format border
fist_cell = 'A1'
last_cell = name_last_clmn + str(row_len_minus1)
rng = sheet.range(fist_cell,last_cell)
for bi in range(7,13):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0xd9d9d9

for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

# Format border of last trow
first_cell = "A" + str(row_len)
last_cell = name_last_clmn + str(row_len)
rng = sheet.range(fist_cell,last_cell)
for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

# Add total for last row
cell = "A" + str(row_len)
rng = sheet.range(cell)
rng.value = "Total"

# Add title
rowl = "{:,d}".format(len(raw_data_bud)+1)
sheet.range("1:2").insert('down')
sheet.range("A1").value = "Budget Data"
sheet.range("A2").value = "Unit: USD " + " " + "UoM:kg"
sheet.range("G2").value = "Length of Rows: " + rowl

sheet.api.Tab.Color = rgb_to_int((239, 209, 241))


# endregion

#region Format Actual Data
# Add new worksheet
sheet2 = wb.sheets.add("Act")
# Hide gridlines for sheet
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Assign values of inventory records to worksheet
sheet2.range("A1").options(index=False).value = raw_data_act
sheet2.autofit(axis="columns")

# Get length of columns, rows
row_len = len(raw_data_act) + 1
row_len_minus1 = row_len - 1

clmn_len = len(raw_data_act.columns)

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

name_last_clmn = excel_column_name(clmn_len)

# Format common features
first_cell = 'A1'
last_cell = name_last_clmn + str(row_len)

rng = sheet2.range(first_cell, last_cell)
rng.row_height = 18
rng.api.Font.Name = 'Verdana'
rng.api.Font.Size = 10
rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
rng.api.WrapText = False

# Format data
first_cell = 'I1'
last_cell = name_last_clmn + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight

first_cell = 'I1'
last_cell = "I" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'J2'
last_cell = "K" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'L2'
last_cell = "M" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'N2'
last_cell = "N" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'O2'
last_cell = "O" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'P2'
last_cell = "P" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'Q2'
last_cell = "Q" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###.####"

first_cell = 'R2'
last_cell = "R" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###"

first_cell = 'T2'
last_cell = "U" + str(row_len)
rng = sheet2.range(first_cell, last_cell)
rng.number_format = "#,###"

# Format first row
first_cell = 'A1'
last_cell = name_last_clmn + "1"
rng = sheet2.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.WrapText = True
rng.color = ("#70d1d4")
rng.api.Font.Color = 0xffffff
rng.api.Font.Bold = True
rng.row_height = 20

# Format border
fist_cell = 'A1'
last_cell = name_last_clmn + str(row_len_minus1)
rng = sheet2.range(fist_cell,last_cell)
for bi in range(7,13):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0xd9d9d9

for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

# Format border of last trow
first_cell = "A" + str(row_len)
last_cell = name_last_clmn + str(row_len)
rng = sheet2.range(fist_cell,last_cell)
for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

# Add total for last row
cell = "A" + str(row_len)
rng = sheet2.range(cell)
rng.value = "Total"

# Add title
rowl = "{:,d}".format(len(raw_data_act)+1)
sheet2.range("1:2").insert('down')
sheet2.range("A1").value = "Actual Data"
sheet2.range("A2").value = "Unit: USD " + " " + "UoM:kg"
sheet2.range("G2").value = "Length of Rows: " + rowl
sheet2.api.Tab.Color = rgb_to_int((112, 209, 212))

#endregion

# region Format Variance Analysis
# Add new sheet
sheet3 = wb.sheets.add("Variance Analysis")

# Hide gridlines for sheet
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

# Assign values of inventory records to worksheet
sheet3.range("A1").options(index=True).value = variance_analysis_ot

sheet3.range("A5").options(index=False).value = variance_factors_amount_ot

sheet3.pictures.add(fig, name='ACES_Plot1', update=True,
                     left=sheet.range('A9').left, top=sheet.range('A9').top)

# Get length of columns, rows
row_len = len(variance_analysis_ot) + 1
clmn_len = len(variance_analysis_ot.columns) + 1

def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ''
    while n > 0:
        n, r = divmod (n - 1, 26)
        name = chr(r + ord('A')) + name
    return name

name_last_clmn = excel_column_name(clmn_len)

# Format common features
first_cell = 'A1'
last_cell = name_last_clmn + str(row_len)

rng = sheet3.range(first_cell, last_cell)
rng.row_height = 18
rng.column_width = 15
rng.api.Font.Name = 'Verdana'
rng.api.Font.Size = 10
rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
rng.api.WrapText = True

#sheet.autofit(axis="columns")

# Format data
first_cell = 'B2'
last_cell = 'G' + str(row_len)
rng = sheet3.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
rng.number_format = "#,###"
#rng.column_width = 15

first_cell = 'H2'
last_cell = 'I' + str(row_len)
rng = sheet3.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
rng.number_format = "#,###.####"

# Format border
fist_cell = 'A1'
last_cell = name_last_clmn + str(row_len)
rng = sheet3.range(fist_cell,last_cell)
for bi in range(7,13):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0xd9d9d9

for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

# Format first row
first_cell = 'A1'
last_cell = name_last_clmn +  "1"
rng = sheet3.range(first_cell, last_cell)
rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000
rng.row_height = 25

# Format last row
first_cell = "A" + str(row_len)
last_cell = name_last_clmn +  str(row_len)
rng = sheet3.range(first_cell, last_cell)
for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

sheet3.range("5:5").value =""
sheet3.range("5:5").insert('down')

sheet3.range("7:7").api.WrapText = True
sheet3.range("7:7").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

rng=sheet3.range("A7","H8")
rng.api.Font.Name = 'Verdana'
rng.api.Font.Size = 10
for bi in range(7,13):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0xd9d9d9

for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

rng=sheet3.range("A7","H7")
for bi in range(7,11):
    rng.api.Borders(bi).Weight = 2
    rng.api.Borders(bi).Color = 0x000000

sheet3.range("A6").value ="Variance Factors and Amounts"
sheet3.range("A6").api.Font.Size = 13
sheet3.range("A6").api.Font.Name = 'Verdana'
sheet3.range("A6").api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft

sheet3.range("A8","H8").number_format = "#,###"

# Add title
# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")

Title = "Operating Income Variance Analysis " "- " + mth_str
sheet3.range("1:4").insert('down')
sheet3.range("A1").value = Title
sheet3.range("A1").api.Font.Size = 14
sheet3.range("A3").value = "The last update time :  " + time_local  +"."
sheet3.range("A4").value = "Analyzed by:  ACES Analytics Team."
sheet3.range("F4").value = "Unit: Revenue/RMC: K'$ | Volume: Ton | Unit Price/RMC: $/kg"
sheet3.range("A1","A4").Fontname = "'Verdana'"
sheet3.api.Tab.Color = rgb_to_int((36, 234, 36))

# Calculate running time to be shown in output file
end = timer()
running_time = "{:,.2f}".format(end - start)
sheet3.range("F3").value = "Running time:  " +  running_time + "  s"

#endregion

"""
## 9. Export worksheets to excel
"""
wb.save(output_file)
wb.close()
app.quit()

# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")

# Print the end for this script
print("The run of script is completed successfully.")
time_local_end = strftime ("%A, %d %b %Y, %H:%M")

#Print running time
running_time2 = "{:,.2f}".format(end - start)
print (running_time2 )

