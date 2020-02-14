import numpy as np
import xlrd
import pandas as  pd
import math
import scipy
from scipy import stats

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

def noisy_elo(year):
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
    loc2 = ("noise.xls")
    wb2 = xlrd.open_workbook(loc2)
    sheet2 = wb2.sheet_by_index(0) 
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
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
        if scorei == 0 and int(sheet2.cell_value(row,2)) == -1:
            scorei = 0
        elif scorej == 0 and int(sheet2.cell_value(row,3)) == -1:
            scorej = 0
        else:
            scorei += int(sheet2.cell_value(row,2))
            scorej += int(sheet2.cell_value(row,3))
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

def noisy_massey(year):
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
    loc2 = ("noise.xls")
    wb2 = xlrd.open_workbook(loc2)
    sheet2 = wb2.sheet_by_index(0) 
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if scorei == 0 and int(sheet2.cell_value(row,2)) == -1:
            scorei = 0
        elif scorej == 0 and int(sheet2.cell_value(row,3)) == -1:
            scorej = 0
        else:
            scorei += int(sheet2.cell_value(row,2))
            scorej += int(sheet2.cell_value(row,3))
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

def noisy_premier(year):
    pl = {}
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    loc2 = ("noise.xls")
    wb2 = xlrd.open_workbook(loc2)
    sheet2 = wb2.sheet_by_index(0) 
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if scorei == 0 and int(sheet2.cell_value(row,2)) == -1:
            scorei = 0
        elif scorej == 0 and int(sheet2.cell_value(row,3)) == -1:
            scorej = 0
        else:
            scorei += int(sheet2.cell_value(row,2))
            scorej += int(sheet2.cell_value(row,3))
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

''' Colley Rating System '''

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

def noisy_colley(year):
    team_w_l = {}
    num_teams = 20
    play_other = 2
    games_team = 38
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    loc2 = ("noise.xls")
    wb2 = xlrd.open_workbook(loc2)
    sheet2 = wb2.sheet_by_index(0) 
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if scorei == 0 and int(sheet2.cell_value(row,2)) == -1:
            scorei = 0
        elif scorej == 0 and int(sheet2.cell_value(row,3)) == -1:
            scorej = 0
        else:
            scorei += int(sheet2.cell_value(row,2))
            scorej += int(sheet2.cell_value(row,3))
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

def main():
    sensitivity_df = pd.DataFrame({"Year": [],  
        "Premier League Sensitivity": [], 
        "Massey Sensitivity": [], 
        "Elo Sensitivity": [],
        "Colley Sensitivity": [],
        "Average Sensitivity": []}
    )
    for year in range(1995,2019):
        # Elo
        elo_rating = list(elo(year).values())
        elo_rating_noisy = list(noisy_elo(year).values())
        # Massey
        massey_rating = list(massey(year).values())
        massey_rating_noisy = list(noisy_massey(year).values())
        # Premier League
        premier_rating = list(premier(year).values())
        premier_rating_noisy = list(noisy_premier(year).values())
        # Colley
        colley_rating = list(colley(year).values())
        colley_rating_noisy = list(noisy_colley(year).values())
        # Sensitivity Table
        df_current = pd.Series({"Year": int(year), 
            "Premier League Sensitivity": scipy.stats.spearmanr(premier_rating,premier_rating_noisy)[0],
            "Massey Sensitivity": scipy.stats.spearmanr(massey_rating,massey_rating_noisy)[0],
            "Elo Sensitivity": scipy.stats.spearmanr(elo_rating,elo_rating_noisy)[0],
            "Colley Sensitivity": scipy.stats.spearmanr(colley_rating,colley_rating_noisy)[0]})
        sensitivity_df = sensitivity_df.append(df_current, ignore_index=True)
        sensitivity_df['Average Sensitivity'] = sensitivity_df[[
            'Premier League Sensitivity',
            'Massey Sensitivity',
            'Elo Sensitivity',
            'Colley Sensitivity']].mean(axis=1)
        sensitivity_df.sort_values(by=['Average Sensitivity'], inplace=True, ascending=False)
    print(sensitivity_df)

if __name__ == "__main__":
    main()
