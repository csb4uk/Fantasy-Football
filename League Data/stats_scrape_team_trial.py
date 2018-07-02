def extract_team_name(team_id_number):
    team_dictionary = {
        1: "The Dictator",
        2: "Octobers Very Own",
        3: "Team Ratliff",
        4: "King Goffrey",
        5: "Pitts Likes Dicks",
        6: "The Last Jedi",
        8: "5 INT 7-8",
        9: "Queen B",
        10: "Team No Trades",
        11: "Team capt queef",
        13: "Team Quintanilla",
        15: "Team Binder",
    }
    print(team_dictionary[team_id_number])
    return team_dictionary[team_id_number]


def debugging():
    open('htmlESPNsoup.html', 'w')
    open('htmlESPNtable.html', 'w')
    open('htmlESPNcategories.html', 'w')
    open('htmlESPNinfo.html', 'w')
    open('htmlESPNplayer_stats.html', 'w')


def parse_team_names(soup, team_id_number):
    for name in soup.find_all("span", attrs={'class', 'alt-info'}):
        name.decompose()
    href_str = "/ffl/clubhouse?leagueId=428408&teamId=" + str(team_id_number) + "&seasonId=2017"
    team_names = soup.find_all('a', attrs={'href': href_str})
    return team_names


def parse_starter_table(soup):
    full_player_table = soup.find('div', attrs={'style': "width: 100%; margin-bottom: 40px; clear: both;"})
    starter_table = full_player_table.find_all('table', attrs={'class': 'playerTableTable tableBody'})
    return starter_table


def parse_bench_table(soup):
    full_player_table = soup.find('div', attrs={'style': "width: 100%; margin-bottom: 40px; clear: both;"})
    bench_table = full_player_table.find_all('table', attrs={'class': 'playerTableTable tableBody hideableGroup'})
    return bench_table


def parse_stat_categories_skill(table):
    category = []
    stats_categories_class = table.find_all('tr', attrs={'class': 'playerTableBgRowSubhead2 tableSubHead'})
    for inst in stats_categories_class:
        stats_categories = inst.find_all('td')
        for categories in stats_categories:
            # if isinstance(categories, Tag):
            file = open('htmlESPNcategories.html', 'a')
            file.write(categories.prettify())
            if categories.get_text():
                category.append(categories.get_text())
    return category


def main():
    from bs4 import BeautifulSoup
    import urllib.request
    from bs4 import Tag
    import xlsxwriter
    import re

    # Ask for input for the week to extract
    # week_number_start = int(input("Enter the Week Number you wish to start extracting data: "))
    # week_number_end = int(input("Enter the Week Number you wish to finish extracting data: "))
    msg = """
    Enter the team that you wish to extract data for:
    \t 1: The Dictator
    \t 2. Octobers Very Own
    \t 3. Team Ratliff
    \t 4. King Goffrey
    \t 5. Pitts Likes Dicks
    \t 6. The Last Jedi
    \t 8. 5 INT 7-8
    \t 9. Queen B
    \t 10. Team No Trades
    \t 11. Team capt queef
    \t 13. Team Quintanilla
    \t 15. Team Binder
    """
    team_id_number = int(input(msg))
    team_name = extract_team_name(team_id_number)

    # Create a new Excel file
    workbook = xlsxwriter.Workbook('C:/Python36/projects/ESPN_Data/Stats/Team Data/' + team_name + '.xlsx')
    # workbook = xlsxwriter.Workbook('C:/Python36/projects/espn/' + team_name + '.xlsx')

    # Initialize variable
    debugging()

    for week in range(1, 2):
        # Name of web page
        wp = "http://games.espn.com/ffl/boxscorefull?leagueId=428408&teamId=" + str(team_id_number) + "&scoringPeriodId=" + str(week) + "&seasonId=2017&view=scoringperiod&version=full"

        # Create html file of webpage
        html_page = urllib.request.urlopen(wp).read()
        soup = BeautifulSoup(html_page, 'lxml')

        # Create the html layout of the ESPN code
        file = open('htmlESPNsoup.html', 'a')
        file.write(soup.prettify())

        # Initialize variable
        excel_row_counter = 0       # Used as a row counter in excel
        excel_column_counter = 0    # Used as a column counter in excel
        worksheet = workbook.add_worksheet("Week " + str(week))

        # Used as a debug to tell you what week the program is on
        print("Week " + str(week))

        # create a soup of the information of the starting lineup
        starter_table = parse_starter_table(soup)
        # create a soup of the information of the bench lineup
        bench_table = parse_bench_table(soup)

        # Create an array of the scoring categories for skill players and defense
        category_skills = parse_stat_categories_skill(starter_table[0])
        category_defense = parse_stat_categories_skill(starter_table[1])
        for table in starter_table:
            if isinstance(table, Tag):
                file = open('htmlESPNtable.html', 'a')
                file.write(table.prettify())
                starter_player_info = table.find_all('tr', attrs={'class': re.compile('pncPlayerRow*')})
                for info in starter_player_info:
                    # if isinstance(info, Tag):
                    file = open('htmlESPNinfo.html', 'a')
                    file.write(info.prettify())
                    player_stats = info.find_all('td')
                    for stats in player_stats:
                        file = open('htmlESPNplayer_stats.html', 'a')
                        file.write(stats.prettify())
                        if stats.get_text():
                            # print(stats.get_text())
                            current_stat = stats.get_text()
                            try:
                                current_stat = int(current_stat)
                                worksheet.write_number(excel_row_counter, excel_column_counter, current_stat)
                            except:
                                try:
                                    current_stat = float(current_stat)
                                    worksheet.write_number(excel_row_counter, excel_column_counter, current_stat)
                                except ValueError:
                                    worksheet.write(excel_row_counter, excel_column_counter, current_stat)
                            excel_column_counter += 1
                    excel_column_counter = 0
                    excel_row_counter += 1
main()
