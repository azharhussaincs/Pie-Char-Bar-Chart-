import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")  # Set interactive backend
import matplotlib.pyplot as plt
import seaborn as sns

# Read Excel file
df = pd.read_excel('NPS CP Data Feb-25.xlsx')

# Process status data
status_array = df["Status"].dropna().values  
status_counts = pd.Series(status_array).value_counts()

# Set dark background for contrast
plt.style.use("dark_background")

# Create a figure with subplots
fig, axes = plt.subplots(1, 2, figsize=(15, 6))  # 1 row, 2 columns

# Define a better color palette
colors = sns.color_palette("plasma", len(status_counts))

# Bar Plot with Outer Borders and Count Labels
bars = sns.barplot(x=status_counts.index, y=status_counts.values, ax=axes[0], hue=status_counts.index, legend=False, palette=colors)
axes[0].set_xlabel("Status", fontsize=12, fontweight="bold", color="white")
axes[0].set_ylabel("Count", fontsize=12, fontweight="bold", color="white")
axes[0].set_title("Distribution of Status Values", fontsize=14, fontweight="bold", color="white")
axes[0].tick_params(axis='x', rotation=45, colors="white")
axes[0].tick_params(axis='y', colors="white")

# Add count labels on bars
for bar in bars.patches:
    height = bar.get_height()
    axes[0].text(bar.get_x() + bar.get_width()/2, height, f'{int(height)}', ha='center', va='bottom', fontsize=12, color='white', fontweight="bold")
    bar.set_edgecolor('black')  # Outer border color
    bar.set_linewidth(1.5)  # Border thickness

# Fix: Manually set labels for the legend
axes[0].legend(handles=bars.patches[:len(status_counts)], labels=status_counts.index.tolist(), title="Status", loc="upper right", fontsize=10, title_fontsize=12)

# Pie Chart with Percentage Labels and Explode Effect
explode_values = [0.05] * len(status_counts)  # Slightly explode all slices
status_counts.plot(
    kind='pie', 
    autopct='%1.1f%%', 
    colors=colors, 
    ax=axes[1], 
    wedgeprops={'edgecolor': 'black', 'linewidth': 1.5}, 
    explode=explode_values
)
axes[1].set_title("Status Distribution", fontsize=14, fontweight="bold", color="white")
axes[1].set_ylabel("")  # Remove default y-label

# Add legend for pie chart
axes[1].legend(labels=status_counts.index.tolist(), title="Status", loc="upper right", fontsize=10, title_fontsize=12)

# Adjust layout for better spacing
plt.tight_layout()

# Show the final visualization
plt.show()
