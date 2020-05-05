import xlrd
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# import xlsm file and get necessary sheets
workbook = xlrd.open_workbook('NBA.xlsm')
for sheet in workbook.sheets():
    if "Trail Blazers" == sheet.name:
        trail = sheet
    if "Warriors" == sheet.name:
        warriors = sheet
        
#setting compare dates
cmp1 = datetime.date(2018, 10, 25)
cmp2 = datetime.date(2018, 11, 20)
cmp3 = datetime.date(2018, 10 ,16)
cmp4 = datetime.date(2018, 11, 10)
trailIdx = []
warriorsIdx = []
first_half_trail = []
first_half_warriors = []
date_trail = []
date_warriors = []

# get data that meet criteria
for i in range(1,trail.nrows):
    s= xlrd.xldate.xldate_as_datetime(trail.cell_value(i,1), workbook.datemode)
    if trail.cell_value(i,2) == "Away" and s.date() >= cmp1 and s.date() <= cmp2: 
        trailIdx.append(i)
        first_half_trail.append(int(trail.cell_value(i,6)))
        print(first_half_trail)
        date_trail.append(s.date())
for i in range(1,warriors.nrows):
    s= xlrd.xldate.xldate_as_datetime(warriors.cell_value(i,1), workbook.datemode)
    if warriors.cell_value(i,2) == "Home" and s.date() >= cmp3 and s.date() <= cmp4: 
        warriorsIdx.append(i)
        first_half_warriors.append(int(warriors.cell_value(i,6)))
        date_warriors.append(s.date())
        
fig = plt.figure(1)

#Table - trail table
trail_table = plt.subplot2grid((3,4), (2,0), colspan=2)
trail_table.table(cellText=[first_half_trail],rowLabels=['Trail Blazers'],colLabels=date_trail)
trail_table.axis("off")

#graph for trail & warriors
trail_graph = plt.subplot2grid((3,4), (0,0), colspan=4,rowspan=2)
trail_graph.set_title("Trail Blazors(Away) vs Warriors(Home)",fontsize=20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d/%Y'))
plt.gca().xaxis.set_major_locator(mdates.DayLocator())
plt.grid(True)
plt.plot(date_trail,first_half_trail,'g-o', markersize=20, label="Trail blazers")
plt.plot(date_warriors,first_half_warriors,'r-o',markersize=20,linestyle='dashed',label="warriors")
plt.xticks(rotation=45)
plt.legend()
trail_graph.yaxis.set_label_coords(0,1.02)
plt.ylabel("Score", rotation=0)
# Table for warriors

warriors_table = plt.subplot2grid((3,4),(2,2), colspan=2)
warriors_table.table(cellText=[first_half_warriors],rowLabels=['Warriors'],colLabels=date_warriors,fontsize=50)
warriors_table.axis("off")

fig.set_size_inches(w=60, h=30)
fig.subplots_adjust(hspace=0.9)

plt.show()