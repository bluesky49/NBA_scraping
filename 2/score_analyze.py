from openpyxl import load_workbook
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

wb = load_workbook('score.xlsx')
ws_pelicans = wb['Pelicans']
ws_lakers = wb['Lakers']


first_half_pelicans = [int(item1.value) + int(item2.value) for item1, item2 in zip(ws_pelicans['E'][1:], ws_pelicans['G'][1:])]
first_half_lakers = [int(item1.value) + int(item2.value) for item1, item2 in zip(ws_lakers['E'][1:], ws_lakers['G'][1:])]
date_lakers = [datetime.datetime.strptime(ws_lakers['B'][index].value, "%B %d, %Y").date() for index in range(1,len(ws_lakers['B']))]
date_pelicans = [datetime.datetime.strptime(ws_pelicans['B'][index].value, "%B %d, %Y").date() for index in range(1,len(ws_pelicans['B']))]


fig = plt.figure(figsize=(10,3))
# #Table - lakers table
lakers_table = plt.subplot2grid((6,4), (2,0), colspan=4,rowspan=2)
lakers_table.table(cellText=[first_half_lakers],colLabels=date_lakers,colWidths=[.048]*len(first_half_lakers))
lakers_table.axis("off")
lakers_table.set_title("Lakers", y=0.1)

# #graph for lakers & pelicans
lakers_graph = plt.subplot2grid((6,4), (0,0), colspan=4,rowspan=2)
lakers_graph.set_title("1st half score of Lakers and Pelicans",fontsize=20)
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%m/%d/%Y'))
plt.gca().xaxis.set_major_locator(mdates.DayLocator())
plt.grid(True)
plt.plot(date_lakers,first_half_lakers,'g-o', markersize=10, label="Lakers")
plt.plot(date_lakers,first_half_pelicans,'r-o',markersize=10,linestyle='dashed',label="Pelicans")
plt.xticks(rotation=45)
plt.legend()
lakers_graph.yaxis.set_label_coords(0,1.02)
plt.ylabel("Score", rotation=0)
# # Table for pelicans

pelicans_table = plt.subplot2grid((6,4),(4,0), colspan=4,rowspan=2)
s = pelicans_table.table(fontsize=10,cellText=[first_half_pelicans],colLabels=date_pelicans, loc='center', colWidths=[.048]*len(first_half_lakers))
pelicans_table.set_title("Pelicans",y=0.7)
pelicans_table.axis("off")

fig.set_size_inches(w=60, h=30)
fig.subplots_adjust(hspace=0.9)

plt.show()