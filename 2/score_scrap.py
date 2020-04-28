from bs4 import BeautifulSoup
import xlsxwriter
import datetime
import requests

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
abb = ["ATL","BKN","BOS","CHA","CHI","CLE","DAL","DEN","DET","GSW","HOU","IND","LAC","LAL","MEM","MIA","MIL","MIN","NOP","NYK","OKC","ORL","PHI","PHX","POR","SAC","SAS","TOR","UTA","WAS"]
# for i in abb:
#     url = "https://www.basketball-reference.com/teams/" + i + "/2020_games.html"
#     res = requests.get(url,headers = headers)
#     soup = BeautifulSoup(res.content, 'html5lib')
url = "https://www.basketball-reference.com/teams/" + "ATL" + "/2020_games.html"
res = requests.get(url,headers = headers)
soup = BeautifulSoup(res.content, 'html5lib')
print(soup.prettify())
    