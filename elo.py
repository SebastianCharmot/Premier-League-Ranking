import numpy as np
import xlrd
import pandas as  pd
import math

num_teams = 20
s_win = 1
s_tie = 0
s_lose = 0.5
K = 10

def pop_teams(year):
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
    elo_rating(team_elo, year)

def mu(rating1, rating2):
    return 1.0 * 1.0 / (1 + 1.0 * math.pow(10, 1.0 * (rating1 - rating2) / 400))

def elo_rating(teams, year):
    loc = ("pl_data/PL_" + str(year) + ".xls")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    rows = sheet.nrows
    for row in range(rows):
        teami = sheet.cell_value(row,0)
        teamj = sheet.cell_value(row,1)
        scorei = int(sheet.cell_value(row,2))
        scorej = int(sheet.cell_value(row,3))
        if scorei > scorej:
            teams[teami] += K * (1 - mu(teams[teamj],teams[teami]))
            teams[teamj] += K * (0 - mu(teams[teami],teams[teamj]))
        elif scorej > scorei:
            teams[teami] += K * (0 - mu(teams[teamj],teams[teami]))
            teams[teamj] += K * (1 - mu(teams[teami],teams[teamj]))
    # print(teams)
    show_elo(teams)

def show_elo(teams):
    print(teams)
    li_teams = []
    elo_teams = []
    for team in teams:
        li_teams.append(team)
        elo_teams.append(teams[team])
    elo_df = pd.DataFrame({"Team": li_teams, "Elo Rating": elo_teams})
    elo_df.sort_values(by=['Elo Rating'], inplace=True, ascending=False)
    print(elo_df.to_string(index=False))

def main():
    pop_teams(2016)

if __name__ == "__main__":
    main()