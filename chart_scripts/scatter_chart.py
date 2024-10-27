import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Load data from the 'Scatter Chart' sheet in ChartData.xlsx
data = pd.read_excel('ChartData.xlsx', sheet_name='Scatter Chart')

# Create folders for saving images
output_dir = 'images/scatter'
os.makedirs(output_dir, exist_ok=True)

### 1. Simple Scatter Chart (January vs. February Sales) ###
plt.figure(figsize=(10, 6))
plt.scatter(data['January'], data['February'], color='b')
for i, employee in enumerate(data['Employee']):
    plt.annotate(employee, (data['January'][i], data['February'][i]), fontsize=8, ha='right')
plt.title('Simple Scatter Chart: January vs February Sales (Each point is an employee)')
plt.xlabel('January Sales')
plt.ylabel('February Sales')
plt.tight_layout()
plt.savefig(f'{output_dir}/simple_scatter_chart.png')
plt.close()

### 2. Scatter Plot Matrix (All Month Pair Comparisons) ###
data_without_employee = data.drop(columns=['Employee'])  # Remove the Employee column for the matrix
sns.pairplot(data_without_employee)
plt.suptitle("Scatter Plot Matrix of Monthly Sales", y=1.02)
plt.savefig(f'{output_dir}/scatter_plot_matrix.png')
plt.close()

### 3. Simple Bubble Chart (April vs May Sales, Bubble size represents March Sales) ###
plt.figure(figsize=(10, 6))
plt.scatter(data['April'], data['May'], s=data['March'] * 10, alpha=0.5, c='g')
for i, employee in enumerate(data['Employee']):
    plt.annotate(employee, (data['April'][i], data['May'][i]), fontsize=8, ha='right')
plt.title('Simple Bubble Chart: April vs May Sales (Bubble size represents March Sales)')
plt.xlabel('April Sales')
plt.ylabel('May Sales')
plt.tight_layout()
plt.savefig(f'{output_dir}/simple_bubble_chart.png')
plt.close()

### 4. Bubble Chart Matrix (Comparing Pairs with Varying Bubble Sizes) ###
# Loop through each consecutive month pair, with the third month as bubble size
for i in range(len(data_without_employee.columns) - 2):
    month_x = data_without_employee.columns[i]
    month_y = data_without_employee.columns[i + 1]
    month_size = data_without_employee.columns[i + 2]

    plt.figure(figsize=(10, 6))
    plt.scatter(data[month_x], data[month_y], s=data[month_size] * 10, alpha=0.5, c='purple')
    for j, employee in enumerate(data['Employee']):
        plt.annotate(employee, (data[month_x][j], data[month_y][j]), fontsize=8, ha='right')
    plt.title(f'Bubble Chart: {month_x} vs {month_y} Sales (Bubble size represents {month_size} Sales)')
    plt.xlabel(f'{month_x} Sales')
    plt.ylabel(f'{month_y} Sales')
    plt.tight_layout()
    plt.savefig(f'{output_dir}/bubble_chart_{month_x}_vs_{month_y}_size_{month_size}.png')
    plt.close()
