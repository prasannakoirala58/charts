import subprocess
import sys

# List of chart script files to run
scripts = [
    'chart_scripts/bar_chart.py',
    'chart_scripts/pie_chart.py',
    'chart_scripts/line_chart.py',
    'chart_scripts/scatter_chart.py',
    'chart_scripts/treemap_chart.py'
]

# Run each script
for script in scripts:
    subprocess.run([sys.executable, script])
