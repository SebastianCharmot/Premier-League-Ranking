import numpy as np
import xlrd
import pandas as  pd

num_teams = 20
# Number of games played by each team
games_team = 38
# How many times each team played every other team
play_other = 2

def pt_differential(year):
    team_pt_diff = {}
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

    # team_pt_diff = sorted(team_pt_diff.keys())
    calc_rating(team_pt_diff)

def calc_rating(team_differential):
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
    teams = sorted(team_differential.keys())
    p = []
    for team in teams:
        p.append(team_differential[team])
    p[-1] = 0
    p = np.matrix(p)
    p = p.T
    #Calculating rating 
    ratings = np.linalg.solve(M,p)
    ratings = ratings.tolist()
    li_ratings = []
    for i in range(num_teams):
        li_ratings.append(round(ratings[i][0],3))
    rating_df = pd.DataFrame({"Team": teams, "Massey Rating": li_ratings})
    rating_df.sort_values(by=['Massey Rating'], inplace=True, ascending=False)
    print(rating_df.to_string(index=False))
    return teams, li_ratings

def calc_d_o():
    pass

def main():
    # Input a year between 1995 and 2018
    pt_differential(2016)

if __name__ == "__main__":
    main()
