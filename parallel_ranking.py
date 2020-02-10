import plotly.graph_objects as go
import pandas as pd

df = pd.read_csv("parallel-data.csv")

fig = go.Figure(data=
    go.Parcoords(
        line = dict(color = "#002366"),
        dimensions = list([
            dict(range = [22,95],
                constraintrange = [4,8],
                label = 'Premier League', values = df['Premier League']),
            dict(range = [-1.1,1.6],
                label = 'Masseys Method', values = df['Massey']),
            dict(range = [1117,1300],
                label = 'Elo', values = df['Elo'])
        ])
    )
)

fig.update_layout(
    plot_bgcolor = 'white',
    paper_bgcolor = 'white'
)

fig.show()