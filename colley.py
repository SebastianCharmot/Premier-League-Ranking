import numpy as np
import xlrd
import pandas as  pd

def colley(year):
    team_w_l = {}
    num_teams = 20
    play_other = 2
    games_team = 38
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if teami not in team_w_l and teamj not in team_w_l:
            team_w_l[teami] = 0
            team_w_l[teamj] = 0
        elif teami not in team_w_l:
            team_w_l[teami] = 0
        elif teamj not in team_w_l:
            team_w_l[teamj] = 0
        if scorei > scorej:
            team_w_l[teami] += 1
            team_w_l[teamj] -= 1
        elif scorej > scorei:
            team_w_l[teami] -= 1
            team_w_l[teamj] += 1
    C = []
    for i in range(num_teams):
        curr_line = []
        for j in range(num_teams):
            curr_line.append(-1*play_other)
        curr_line[i] = 2+games_team
        C.append(curr_line)
    C = np.matrix(C)
    teams = sorted(team_w_l.keys())
    p = []
    for team in teams:
        p.append(team_w_l[team])
    p = np.matrix(p)
    p = p.T
    #Calculating Colley rating 
    ratings = np.linalg.solve(C,p)
    ratings = ratings.tolist()
    li_ratings = []
    for i in range(num_teams):
        li_ratings.append(round(ratings[i][0],3))
    colley_rating = {}
    for i in range(num_teams):
        colley_rating[teams[i]] = li_ratings[i]
    return colley_rating

colley(2016)
    # rating_df = pd.DataFrame({"Team": teams, "Massey Rating": li_ratings})
    # rating_df.sort_values(by=['Massey Rating'], inplace=True, ascending=False)
    # print(rating_df.to_string(index=False))
    # return teams, li_ratings
