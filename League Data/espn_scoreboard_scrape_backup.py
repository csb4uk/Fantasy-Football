def week_extract():
    week_number = input("Enter the Week Number you wish to extract data: ")
    return week_number


def debugging():
    open('htmlESPNdiv.html', 'w')
    open('htmlESPNtable.html', 'w')
    open('htmlESPNmatchupinfo.html', 'w')
    open('htmlESPNname.html', 'w')
    open('htmlESPNscore.html', 'w')


def parse_divs(soup):
    divs = soup.find('div', attrs={'id': 'scoreboardMatchups'})
    return divs


def parse_table(div_soup):
    tables = div_soup.find_all('table', attrs={'class': 'ptsBased matchup'})
    return tables


def parse_tr(table_soup):
    team_info = table_soup.find_all('tr')
    return team_info


def parse_team_info(matchup_soup):
    team_info = matchup_soup.find_all('a')
    return team_info


def parse_score_info(matchup_soup):
    score_info = matchup_soup.find_all('td', attrs={'class': 'score'})
    return score_info


def column_widths_A(current_team, max_str_length_team):
    # Figure out the column width of Column A
    if len(str(current_team)) > len(str(max_str_length_team)):
        max_str_length_team = len(current_team)
    return max_str_length_team


def column_widths_B(current_team_score, max_str_length_score):
    # Figure out the column width of Column B
    if len(str(current_team_score)) > len(str(max_str_length_score)):
        max_str_length_score = len(current_team_score)
    return max_str_length_score


def main():
    from bs4 import BeautifulSoup
    import urllib.request
    from bs4 import Tag
    import xlsxwriter

    # Ask for input for the week to extract
    week_number = week_extract()
    week_str = "Week " + str(week_number)

    # Create a new Excel file
    workbook = xlsxwriter.Workbook('C:/Python36/projects/ESPN_Data/Scoreboard_Info/' + str(week_str) + '.xlsx')
    worksheet = workbook.add_worksheet(week_str)

    # Name of web page
    wp = "http://games.espn.com/ffl/scoreboard?leagueId=428408&matchupPeriodId=" + str(week_number)

    # Create html file of webpage
    html_page = urllib.request.urlopen(wp).read()
    soup = BeautifulSoup(html_page, 'html.parser')

    # Initialize variable
    matchup_counter = 1         # Used as a counter for the matchups
    excel_row_counter = 0       # Used as a row counter in excel
    excel_column_counter = 0    # Used as a column counter in excel
    max_str_length_team = 0     # Used to size the team name column at the end of the program
    max_str_length_score = 0    # Used to size the score column at the end of the program
    # debugging()

    # Find the location where all matchups are kept
    divs = parse_divs(soup)
    for div in divs:
        if isinstance(div, Tag):
            # file = open('htmlESPNdiv.html', 'a')
            # file.write(div.prettify())
            # Find ptsBased matchup to locate the matchup between the two teams.  There should be 6 instances
            tables = parse_table(div)
            for table in tables:
                if isinstance(table, Tag):
                    # file = open('htmlESPNtable.html', 'a')
                    # file.write(table.prettify())
                    matchup_number = 'Matchup ' + str(matchup_counter)
                    worksheet.write(excel_row_counter, excel_column_counter, matchup_number)

                    # Increment counters for the matchup and excel row
                    matchup_counter += 1
                    excel_row_counter += 1

                    # Separate each team from the matchup
                    matchup_info = parse_tr(table)
                    for matchup in matchup_info:
                        if isinstance(matchup, Tag):
                            # file = open('htmlESPNmatchupinfo.html', 'a')
                            # file.write(matchup.prettify())
                            try:
                                # Find the team info in the <a> block
                                team_info = parse_team_info(matchup)
                                # Find score info in the <td class = "score"> block
                                score_info = parse_score_info(matchup)
                                for t_name in team_info:
                                    # file = open('htmlESPNname.html', 'a')
                                    # file.write(t_name.prettify())
                                    # Only extract the team name
                                    if t_name.has_attr('title'):
                                        for score in score_info:
                                            # file = open('htmlESPNscore.html', 'a')
                                            # file.write(score.prettify())

                                            # Assign team name to variable
                                            current_team = t_name.get_text()
                                            max_str_length_team = column_widths_A(current_team, max_str_length_team)
                                            current_team_score = score.get_text()
                                            max_str_length_score = column_widths_B(current_team_score, max_str_length_score)

                                            # Debug print names and scores
                                            # print(current_team + "\n\t" + current_team_score)

                                            # Write the team name to Column A
                                            worksheet.write(excel_row_counter, excel_column_counter, current_team)
                                            # Write the team score to Column B
                                            worksheet.write(excel_row_counter, excel_column_counter + 1, current_team_score)
                                            # Increment row coutner
                                            excel_row_counter += 1
                            except:
                                print('bad string')
                                continue
                    excel_row_counter += 1
    worksheet.set_column('A:A', max_str_length_team + 5)
    worksheet.set_column('B:B', max_str_length_score + 5)
main()
