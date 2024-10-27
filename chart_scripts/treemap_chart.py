import pandas as pd
import matplotlib.pyplot as plt
import squarify
import os

# Load data from the 'Treemap' sheet in ChartData.xlsx, specifying that it has no headers
data = pd.read_excel('ChartData.xlsx', sheet_name='Treemap', header=None, names=['Company', 'Spending'])

# Create the folder for treemap images if it doesn't exist
output_dir = 'images/treemap'
os.makedirs(output_dir, exist_ok=True)

# Treemap Chart
plt.figure(figsize=(10, 6))
squarify.plot(sizes=data['Spending'], label=data['Company'], alpha=0.8)
plt.title('Treemap Chart of Company Spending')
plt.axis('off')

# Save the chart
plt.savefig(f'{output_dir}/treemap_chart.png')
plt.close()
