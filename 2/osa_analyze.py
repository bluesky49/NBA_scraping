from openpyxl import load_workbook
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

wb = load_workbook('score.xlsx')
ws_pelicans = wb['Pelicans']
ws_lakers = wb['Lakers']

away_lakers = [ws_lakers['M'][i].value for i in range(len(ws_lakers['M'])) if ws_lakers['C'][i].value == "Away"]
home_lakers = [ws_lakers['M'][i].value for i in range(len(ws_lakers['M'])) if ws_lakers['C'][i].value == "Home"]
away_pelicans = [ws_pelicans['M'][i].value for i in range(len(ws_pelicans['M'])) if ws_pelicans['C'][i].value == "Away"]
home_pelicans = [ws_pelicans['M'][i].value for i in range(len(ws_pelicans['M'])) if ws_pelicans['C'][i].value == "Home"]


away_average_lakers = []
away_average_pelicans = []
home_average_lakers = []
home_average_pelicans = []
for i in range(len(away_lakers)):
    sum = 0
    for j in range(i+1):
        sum += int(away_lakers[j])
    away_average_lakers.append(round(sum/(i+1),2))
for i in range(len(away_pelicans)):
    sum = 0
    for j in range(i+1):
        sum += int(away_pelicans[j])
    away_average_pelicans.append(round(sum/(i+1),2))
for i in range(len(home_lakers)):
    sum = 0
    for j in range(i+1):
        sum += int(home_lakers[j])
    home_average_lakers.append(round(sum/(i+1),2))
for i in range(len(home_pelicans)):
    sum = 0
    for j in range(i+1):
        sum += int(home_pelicans[j])
    home_average_pelicans.append(round(sum/(i+1),2))
    
marginal_lakers = [int(away_lakers[i]) for i in range(len(away_lakers))]
    
lakers_graph = plt.subplot2grid((6,4), (0,0), colspan=4,rowspan=2)
lakers_graph.set_title("Average away score and Marginal score of Lakers",fontsize=20)
plt.grid(True)
x = range(1,len(away_average_lakers)+1)
plt.plot(x,away_average_lakers,'g-o', markersize=10, label="Average score")
plt.plot(range(1,len(marginal_lakers)+1),marginal_lakers,'r-o',markersize=10,linestyle='dashed',label="Margine Score")
plt.legend()


lakers_table = plt.subplot2grid((6,4), (2,0), colspan=4,rowspan=2)
s= lakers_table.table(cellText=[away_average_lakers],colLabels=x,colWidths=[.048]*len(away_average_lakers))
lakers_table.axis("off")
lakers_table.set_title("Average away score of Lakers by game numbers", y=0.1)
s.auto_set_font_size(False)
s.set_fontsize(12)
s.scale(1, 2)

margine_table = plt.subplot2grid((6,4), (4,0), colspan=4,rowspan=2)
s= margine_table.table(cellText=[marginal_lakers],colLabels=x,colWidths=[.048]*len(marginal_lakers))
margine_table.axis("off")
margine_table.set_title("Margine", y=0.1)
s.auto_set_font_size(False)
s.set_fontsize(12)
s.scale(1, 2)

# pelicans_graph = plt.subplot2grid((4,4), (2,0), colspan=4,rowspan=2)
# pelicans_graph.set_title("Average away score of Pelicans",fontsize=20)
# plt.grid(True)
# x = range(len(away_average_pelicans))
# plt.plot(range(len(away_average_pelicans)),away_average_pelicans,'r-o',markersize=10,linestyle='dashed')



plt.legend()
fig = plt.figure(1)
fig.set_size_inches(w=60, h=30)
fig.subplots_adjust(hspace=0.9)

plt.show()