from openpyxl import load_workbook
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from numpy import dot,array,empty_like
from matplotlib.path import Path

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
    
marginal_lakers = [int(home_lakers[i]) for i in range(len(home_lakers))]

def make_path(x1,y1,x2,y2):
    return Path([[x1,y1],[x1,y2],[x2,y2],[x2,y1]])

def perp( a ) :
    b = empty_like(a)
    b[0] = -a[1]
    b[1] = a[0]
    return b


def seg_intersect(a1,a2, b1,b2) :
    da = a2-a1
    db = b2-b1
    dp = a1-b1
    dap = perp(da)
    denom = dot( dap, db)
    num = dot( dap, dp )
   
    x3 = ((num / denom.astype(float))*db + b1)[0]
    y3 = ((num / denom.astype(float))*db + b1)[1]
    p1 = make_path(a1[0],a1[1],a2[0],a2[1])
    p2 = make_path(b1[0],b1[1],b2[0],b2[1])
    if p1.contains_point([x3,y3]) and p2.contains_point([x3,y3]):
        return round(x3,2), round(y3,2)
    else:
        return False

optimum_score_x = []
optimum_score_y = []
x = range(1,len(home_average_lakers)+1)
for i in range(len(x)-1):
    p1 = array([x[i],home_average_lakers[i]])
    p2 = array([x[i+1],home_average_lakers[i+1]])
    p3 = array([x[i],marginal_lakers[i]])
    p4 = array([x[i+1],marginal_lakers[i+1]])
    # print (seg_intersect( p1,p2, p3,p4))
    if seg_intersect( p1,p2, p3,p4):
        optimum_score_x.append(seg_intersect(p1,p2,p3,p4)[0])
        optimum_score_y.append(seg_intersect(p1,p2,p3,p4)[1])
print(optimum_score_x, optimum_score_y)
  
lakers_graph = plt.subplot2grid((8,4), (0,0), colspan=4,rowspan=2)
lakers_graph.set_title("Average away score and Marginal score of Lakers",fontsize=20)
plt.grid(True)

plt.plot(x,home_average_lakers,'g-o', markersize=10, label="Average score")
plt.plot(range(1,len(marginal_lakers)+1),marginal_lakers,'r-o',markersize=10,linestyle='dashed',label="Margine Score")
plt.plot(optimum_score_x,optimum_score_y,'bo',label="Optimum score")
plt.legend()


lakers_table = plt.subplot2grid((8,4), (2,0), colspan=4,rowspan=2)
s= lakers_table.table(cellText=[home_average_lakers],colLabels=x,colWidths=[.048]*len(home_average_lakers))
lakers_table.axis("off")
lakers_table.set_title("Average home score of Lakers by game numbers", y=0.1)
s.auto_set_font_size(False)
s.set_fontsize(12)
s.scale(1, 1)

margine_table = plt.subplot2grid((8,4), (4,0), colspan=4,rowspan=2)
s= margine_table.table(cellText=[marginal_lakers],colLabels=x,colWidths=[.048]*len(marginal_lakers))
margine_table.axis("off")
margine_table.set_title("Margine", y=0.1)
s.auto_set_font_size(False)
s.set_fontsize(12)
s.scale(1,1)

optimum_table = plt.subplot2grid((8,4), (6,0), colspan=6,rowspan=2)
s= optimum_table.table(cellText=[optimum_score_y],colLabels=optimum_score_x,colWidths=[.048]*len(optimum_score_x))
optimum_table.axis("off")
optimum_table.set_title("Optimum", y=0.3)
s.auto_set_font_size(False)
s.set_fontsize(12)
s.scale(1,1)

plt.legend()
fig = plt.figure(1)
fig.set_size_inches(w=60, h=30)
fig.subplots_adjust(hspace=0.9)

plt.show()
