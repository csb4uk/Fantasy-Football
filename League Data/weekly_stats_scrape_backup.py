def week_extract():
    week_number = input("Enter the Week Number you wish to extract data: ")
    return week_number


def debugging():
    open('htmlESPNtable.html', 'w')
    open('htmlESPNstat.html', 'w')


def parse_team_names(soup, team_id_number):
    for name in soup.find_all("span", attrs={'class', 'alt-info'}):
        name.decompose()
    href_str = "/ffl/clubhouse?leagueId=428408&teamId=" + str(team_id_number) + "&seasonId=2017"
    team_names = soup.find_all('a', attrs={'href': href_str})
    return team_names


def parse_starter_table(soup):
    starter_table = soup.find_all('table', attrs={'class': ['playerTableTable tableBody', "playerTableTable tableBody hideableGroup"], 'id': ['playertable_0', 'playertable_1', 'playertable_2']})
    return starter_table


def parse_player_stats(table):
    player_stats = table.find_all('tr', attrs={'class': ['pncPlayerRow playerTableBgRow0', 'pncPlayerRow playerTableBgRow1']})
    return player_stats


def parse_stat_categories(table):
    stats_categories_class = table.find_all('tr', attrs={'class': 'playerTableBgRowSubhead2 tableSubHead'})
    for inst in stats_categories_class:
        stats_categories = inst.find_all('td')
    return stats_categories


def main():
    from bs4 import BeautifulSoup
    import urllib.request
    from bs4 import Tag
    import xlsxwriter

    # Ask for input for the week to extract
    week_number = week_extract()
    week_str = "Week " + str(week_number)

    # Create a new Excel file
    workbook = xlsxwriter.Workbook('C:/Python36/projects/ESPN_Data/Stats/Weekly Data/' + str(week_str) + '.xlsx')
    # workbook = xlsxwriter.Workbook('C:/Python36/projects/espn/' + str(week_str) + '.xlsx')

    # Initialize variable
    team_counter = 0
    # debugging()

    for team_id_number in range(1, 16):
        # Name of web page
        wp = "http://games.espn.com/ffl/boxscorefull?leagueId=428408&teamId=" + str(team_id_number) + "&scoringPeriodId=" + str(week_number) + "&seasonId=2017&view=scoringperiod&version=full"

        # Create html file of webpage
        html_page = urllib.request.urlopen(wp).read()
        soup = BeautifulSoup(html_page, 'lxml')

        # file = open('htmlESPNsoup.html', 'a')
        # file.write(soup.prettify())

        # Initialize variable
        excel_row_counter = 0       # Used as a row counter in excel
        excel_column_counter = 0    # Used as a column counter in excel
        max_str_col_a = 0     # Used to size the team name column at the end of the program
        max_str_length_score = 0    # Used to size the score column at the end of the program

        team_names = parse_team_names(soup, team_id_number)
        for team_name in team_names:
            if isinstance(team_name, Tag) and team_name.string != "":
                team_counter += 1
                worksheet = workbook.add_worksheet(team_name.string)
                print("\n\n" + team_name.string)
                print("Team :" + str(team_counter))
                worksheet.write(excel_row_counter, excel_column_counter, team_name.string)
                excel_row_counter += 1
                starter_table = parse_starter_table(soup)
                for table in starter_table:
                    if isinstance(table, Tag):
                        # file = open('htmlESPNtable.html', 'a')
                        # file.write(table.prettify())
                        stats_categories = parse_stat_categories(table)
                        for categories in stats_categories:
                            # if isinstance(categories, Tag):
                            # file = open('htmlESPNcategories.html', 'a')
                            # file.write(categories.prettify())
                            if categories.get_text():
                                worksheet.write(excel_row_counter, excel_column_counter, categories.get_text())
                                str_len = len(categories.get_text()) + 5
                                worksheet.set_column(excel_column_counter, excel_column_counter, str_len)
                                excel_column_counter += 1
                        excel_row_counter += 1
                        excel_column_counter = 0
                        player_info = table.find_all('tr', attrs={'class': ['pncPlayerRow playerTableBgRow0', 'pncPlayerRow playerTableBgRow1']})
                        for info in player_info:
                            # if isinstance(info, Tag):
                            file = open('htmlESPNinfo.html', 'a')
                            file.write(info.prettify())
                            player_stats = info.find_all('td')
                            for stats in player_stats:
                                file = open('htmlESPNplayer_stats.html', 'a')
                                file.write(stats.prettify())
                                if stats.get_text():
                                    # print(stats.get_text())
                                    worksheet.write(excel_row_counter, excel_column_counter, stats.get_text())
                                    excel_column_counter += 1
                            excel_column_counter = 0
                            excel_row_counter += 1
main()
