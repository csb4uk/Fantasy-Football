from bs4 import BeautifulSoup
import urllib.request
import requests
import csv

# output file name
file_name_1 = 'C:/Python36/projects/ESPN_Data/Projections/projectionsESPN.csv'
# open the file name and write to it
f = csv.writer(open(file_name_1, 'w', newline=''))

# create an error handling file
file_name_2 = 'C:/Python36/projects/ESPN_Data/Projections/errorESPN.csv'
errorFile = csv.writer(open(file_name_2, 'w', newline=''))

week_number = int(input("\nEnter the week you would like to extract projections: "))
number_of_players = int(input("Enter the number of players you would like to extract: "))

# name of the webpage
wp_1 = "http://games.espn.com/ffl/tools/projections?leagueId=428408&startIndex="
x = 0   # initialize counter
wp_2 = "&scoringPeriodId="
wp_3 = "&seasonId=2017"

# # Create a file that prints out the ESPN html code
# file = open('htmlESPN.txt', 'w')
# file.write(soup.prettify())

while (x < number_of_players):
    # Create an html version of the webpage
    html_page = urllib.request.urlopen(wp_1 + str(x) + wp_2 + str(week_number) + wp_3).read()

    # Create soup
    soup = BeautifulSoup(html_page, "html.parser")
# tables = soup.find_all('table')
# for table in tables:
    # file = open('htmlESPN.html', 'w')
    # file.write(table.prettify())
    tableStats = soup.find("table", {"class": "playerTableTable tableBody"})
    trs = tableStats.findAll('tr')
    for tr in trs:
        tds = tr.findAll('td')
        try:
            name = str(tds[0].get_text().strip())
            points = str(tds[15].get_text().strip())
            f.writerow([name,points])

        except:
            print('bad string')
            continue
    x += 40
