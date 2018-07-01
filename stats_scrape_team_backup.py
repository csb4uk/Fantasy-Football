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
    open('htmlESPNcategories.html', 'w')
    open('htmlESPNplayer_table.html', 'w')
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
            # file = open('htmlESPNcategories.html', 'a')
            # file.write(categories.prettify())
            if categories.get_text():
                category.append(categories.get_text())
    return category


def parse_player_data(player_table):
    from bs4 import Tag
    import re
    array_stat = []
    player_stat_array = []
    for each_table in player_table:
        if isinstance(each_table, Tag):
            # file = open('htmlESPNplayer_table.html', 'a')
            # file.write(each_table.prettify())
            all_player_info = each_table.find_all('tr', attrs={'class': re.compile('pncPlayerRow*')})
            for each_player_info in all_player_info:
                # if isinstance(info, Tag):
                # file = open('htmlESPNinfo.html', 'a')
                # file.write(each_player_info.prettify())
                player_stats = each_player_info.find_all('td')
                for stats in player_stats:
                    # if isinstance(info, Tag):
                    #     file = open('htmlESPNplayer_stats.html', 'a')
                    #     file.write(stats.prettify())
                    try:
                        if stats.get_text():
                            # print(stats.get_text())
                            array_stat.append(stats.get_text())
                            # print(array_stat)
                    except:
                        continue
                player_stat_array.append(array_stat)
                array_stat = []
    return player_stat_array


def main():
    from bs4 import BeautifulSoup
    import urllib.request
    import xlsxwriter

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
    workbook = xlsxwriter.Workbook('C:/Python36/projects/ESPN_Data/Stats/Team Data/Excel/' + team_name + '.xlsx')
    # workbook = xlsxwriter.Workbook('C:/Python36/projects/espn/' + team_name + '.xlsx')

    # Initialize variable
    # debugging()
    defense_array = []

    for week in range(1, 17):
        # Name of web page
        wp = "http://games.espn.com/ffl/boxscorefull?leagueId=428408&teamId=" + str(team_id_number) + "&scoringPeriodId=" + str(week) + "&seasonId=2017&view=scoringperiod&version=full"

        # Create html file of webpage
        html_page = urllib.request.urlopen(wp).read()
        soup = BeautifulSoup(html_page, 'lxml')

        # Create the html layout of the ESPN code
        # file = open('htmlESPNsoup.html', 'a')
        # file.write(soup.prettify())

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
        starter_data = parse_player_data(starter_table)
        bench_data = parse_player_data(bench_table)

        for skill in range(len(category_skills)):
            worksheet.write(excel_row_counter, excel_column_counter, category_skills[skill])
            excel_column_counter += 1

        excel_row_counter += 1
        excel_column_counter = 0

        for data in range(len(starter_data)):
            if len(starter_data[data]) > 12:
                for each_stat in starter_data[data]:
                        try:
                            each_stat = int(each_stat)
                            worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                        except:
                            try:
                                each_stat = float(each_stat)
                                worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                            except ValueError:
                                worksheet.write(excel_row_counter, excel_column_counter, each_stat)
                        excel_column_counter += 1
                excel_row_counter += 1
                excel_column_counter = 0
            else:
                # print(starter_data[data])
                defense_array.append(starter_data[data])

        for data in range(len(bench_data)):
            if len(bench_data[data]) > 12:
                for each_stat in bench_data[data]:
                    try:
                        each_stat = int(each_stat)
                        worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                    except:
                        try:
                            each_stat = float(each_stat)
                            worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                        except ValueError:
                            worksheet.write(excel_row_counter, excel_column_counter, each_stat)
                    excel_column_counter += 1
                excel_row_counter += 1
                excel_column_counter = 0
            else:
                # print(bench_data[data])
                defense_array.append(bench_data[data])


        for skill in range(len(category_defense)):
            worksheet.write(excel_row_counter, excel_column_counter, category_defense[skill])
            excel_column_counter += 1
        excel_row_counter += 1
        excel_column_counter = 0

        for defense in range(len(defense_array)):
            for each_stat in defense_array[defense]:
                try:
                    each_stat = int(each_stat)
                    worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                except:
                    try:
                        each_stat = float(each_stat)
                        worksheet.write_number(excel_row_counter, excel_column_counter, each_stat)
                    except ValueError:
                        worksheet.write(excel_row_counter, excel_column_counter, each_stat)
                excel_column_counter += 1
            excel_row_counter += 1
            excel_column_counter = 0

main()
