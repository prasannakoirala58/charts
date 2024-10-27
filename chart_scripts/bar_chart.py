import pandas as pd
import matplotlib.pyplot as plt
import os

# Load data from the 'Bar Chart' sheet in ChartData.xlsx
data = pd.read_excel('ChartData.xlsx', sheet_name='Bar Chart')

# Create the folder for bar chart images if it doesn't exist
output_dir = 'images/bar'
os.makedirs(output_dir, exist_ok=True)

# Vertical Bar Chart for Quarterly Sales by Employee
plt.figure(figsize=(10, 6))
data.plot(kind='bar', x='Employee', y=['Q1 Sales', 'Q2 Sales', 'Q3 Sales', 'Q4 Sales'])
plt.title('Quarterly Sales by Employee')
plt.xlabel('Employee')
plt.ylabel('Sales')
plt.legend(title='Quarter')
plt.tight_layout()
plt.savefig(f'{output_dir}/bar_chart.png')  # Saves the vertical bar chart in the bar folder
plt.close()

# Horizontal Bar Chart for Quarterly Sales by Employee
plt.figure(figsize=(10, 6))
data.plot(kind='barh', x='Employee', y=['Q1 Sales', 'Q2 Sales', 'Q3 Sales', 'Q4 Sales'])
plt.title('Quarterly Sales by Employee (Horizontal)')
plt.xlabel('Sales')
plt.ylabel('Employee')
plt.legend(title='Quarter')
plt.tight_layout()
plt.savefig(f'{output_dir}/column_chart.png')  # Saves the horizontal bar chart in the bar folder
plt.close()
