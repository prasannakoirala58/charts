import pandas as pd
import matplotlib.pyplot as plt
import os

# Load data from the 'Line Chart' sheet in ChartData.xlsx
data = pd.read_excel('ChartData.xlsx', sheet_name='Line Chart')

# Set the index to 'Employee' so each employee's data is in one row
data.set_index('Employee', inplace=True)

# Create the folders if they don't exist
os.makedirs('images/line/individual', exist_ok=True)  # Folder for individual employee charts
os.makedirs('images/line', exist_ok=True)  # Folder for multiline chart

# Line Chart for each employee (saves each employee's line chart in the individual folder)
for employee in data.index:
    plt.figure(figsize=(8, 5))
    plt.plot(data.columns, data.loc[employee], marker='o', linestyle='-', label=employee)
    plt.title(f'Monthly Sales for {employee}')
    plt.xlabel('Month')
    plt.ylabel('Sales')
    plt.tight_layout()
    plt.savefig(f'images/line/individual/line_chart_{employee}.png')  # Saves each employee's line chart separately
    plt.close()  # Closes the figure after saving

# Multiline Chart (all employees on the same chart)
plt.figure(figsize=(10, 6))
for employee in data.index:
    plt.plot(data.columns, data.loc[employee], marker='o', linestyle='-', label=employee)

plt.title('Monthly Sales by Employee (Multiline Chart)')
plt.xlabel('Month')
plt.ylabel('Sales')
plt.legend(title='Employee')
plt.tight_layout()
plt.savefig('images/line/multiline_chart.png')  # Saves the multiline chart in the line folder
plt.close()
