import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import os

# Load data from the 'Pie Chart' sheet in ChartData.xlsx
data = pd.read_excel('ChartData.xlsx', sheet_name='Pie Chart')

# Create the folder for pie chart images if it doesn't exist
output_dir = 'images/pie'
os.makedirs(output_dir, exist_ok=True)

# 2D Pie Chart (Matplotlib)
plt.figure()
plt.pie(data['Dollars Spent'], labels=data['Company'], autopct='%1.1f%%')
plt.title('2D Pie Chart')
plt.savefig(f'{output_dir}/pie_chart_2d.png')  # Saves the 2D pie chart in the pie folder
plt.close()

# Donut Chart (Matplotlib)
plt.figure()
wedges, texts, autotexts = plt.pie(data['Dollars Spent'], labels=data['Company'], autopct='%1.1f%%')
plt.gca().add_artist(plt.Circle((0, 0), 0.70, fc='white'))  # Adding a white circle for the donut effect
plt.title('Donut Chart')
plt.savefig(f'{output_dir}/donut_chart.png')  # Saves the donut chart in the pie folder
plt.close()

# 3D Pie Chart (Plotly)
fig = go.Figure(go.Pie(
    labels=data['Company'],
    values=data['Dollars Spent'],
    title='3D Pie Chart',
    pull=[0.1 if i == 0 else 0 for i in range(len(data))]  # Pulling the first slice slightly for effect
))

fig.update_traces(textinfo='percent+label', textfont_size=15, marker=dict(line=dict(color='#000000', width=2)))
fig.write_image(f"{output_dir}/pie_chart_3d.png")  # Saves the 3D pie chart in the pie folder
