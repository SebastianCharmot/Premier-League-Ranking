import numpy as np
import xlrd
import pandas as  pd
import math
import plotly.graph_objects as go


''' Elo Rating System '''

def mu(rating1, rating2):
        return 1.0 * 1.0 / (1 + 1.0 * math.pow(10, 1.0 * (rating1 - rating2) / 400))

def elo(year):
    s_win = 1
    s_tie = 0
    s_lose = 0.5
    K = 10
    num_teams = 20
    team_elo = {}
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if teami not in team_elo and teamj not in team_elo:
            team_elo[teami] = 1200
            team_elo[teamj] = 1200
        elif teami not in team_elo:
            team_elo[teami] = 1200
        else:
            team_elo[teamj] = 1200
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if scorei > scorej:
            team_elo[teami] += K * (1 - mu(team_elo[teamj],team_elo[teami]))
            team_elo[teamj] += K * (0 - mu(team_elo[teami],team_elo[teamj]))
        elif scorej > scorei:
            team_elo[teami] += K * (0 - mu(team_elo[teamj],team_elo[teami]))
            team_elo[teamj] += K * (1 - mu(team_elo[teami],team_elo[teamj]))
    return team_elo

''' Masseys Rating System '''

def massey(year):
    num_teams = 20
    # Number of games played by each team
    games_team = 38
    # How many times each team played every other team
    play_other = 2
    team_pt_diff = {}
    massey_ranking = {}
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if teami not in team_pt_diff and teamj not in team_pt_diff:
            team_pt_diff[teami] = scorei - scorej
            team_pt_diff[teamj] = scorej - scorei
        elif teami not in team_pt_diff:
            team_pt_diff[teami] = scorei - scorej
            team_pt_diff[teamj] += scorej - scorei
        elif teamj not in team_pt_diff:
            team_pt_diff[teami] += scorei - scorej
            team_pt_diff[teamj] = scorej - scorei
        else:
            team_pt_diff[teami] += scorei - scorej
            team_pt_diff[teamj] += scorej - scorei
    M = []
    last_row = []
    for i in range(num_teams-1):
        curr_line = []
        for j in range(num_teams):
            curr_line.append(-1*play_other)
        curr_line[i] = games_team
        M.append(curr_line)
        last_row.append(1)
    # Massey's adjustment to last row 
    last_row.append(1)
    M.append(last_row)
    M = np.matrix(M)
    # Cleaning team matrix 
    teams = sorted(team_pt_diff.keys())
    p = []
    for team in teams:
        p.append(team_pt_diff[team])
    p[-1] = 0
    p = np.matrix(p)
    p = p.T
    #Calculating rating 
    ratings = np.linalg.solve(M,p)
    ratings = ratings.tolist()
    for i in range(num_teams):
        massey_ranking[teams[i]] = round(ratings[i][0],3)
    return massey_ranking

''' Premier League Rating System '''

def premier(year):
    pl = {}
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if teami not in pl:
            pl[teami] = 0
        if teamj not in pl:
            pl[teamj] = 0
        if scorei == scorej:
            pl[teami] += 1
            pl[teamj] += 1
        elif scorei > scorej:
            pl[teami] += 3
        else:
            pl[teamj] += 3
    return pl

''' Parallel Plot'''

def main():
    year = 2014
    elo_rating = elo(year)
    massey_rating = massey(year)
    premier_rating = premier(year)
    teams = elo_rating.keys()
    ranking_df = pd.DataFrame({"Team": [],"Premier League Rating": [], "Massey Rating": [], "Elo Rating": []})
    for team in teams:
        df_current = pd.Series({"Team": team,"Premier League Rating": premier_rating[team], "Massey Rating": massey_rating[team], "Elo Rating": elo_rating[team]},index=None)
        ranking_df = ranking_df.append(df_current, ignore_index=True)
    # Parallel Plot
    fig = go.Figure(data=
    go.Parcoords(
        line = dict(color = "#002366"),
        dimensions = list([
            dict(range = [22,95],
                constraintrange = [4,8],
                label = 'Premier League', values = ranking_df['Premier League Rating']),
            dict(range = [-1.1,1.6],
                label = 'Masseys Method', values = ranking_df['Massey Rating']),
            dict(range = [1117,1300],
                label = 'Elo', values = ranking_df['Elo Rating'])
            ])
        )
    )
    fig.update_layout(
        plot_bgcolor = 'red',
        paper_bgcolor = 'white'
        )
    fig.show()

if __name__ == "__main__":
    main()
