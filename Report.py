import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")  # Set interactive backend
import matplotlib.pyplot as plt
import seaborn as sns

# Read Excel file
df = pd.read_excel('NPS CP Data Feb-25.xlsx')

# Process status data
status_counts = df["Status"].value_counts()

# Set white background
plt.style.use("default")  
sns.set_style("white")  # Remove background grid lines

# Create figure with subplots
fig, axes = plt.subplots(1, 3, figsize=(22, 6))  

# Set the main title for the figure
fig.suptitle("This is system generated report", fontsize=16, fontweight="bold", color="black")

# Define custom color mapping
status_colors = {"Pending": "blue", "Off Grid": "red", "Closed": "green"}

# Apply colors to status categories
colors = [status_colors.get(status, "gray") for status in status_counts.index]

# Bar Plot (Status Count)
bars = sns.barplot(x=status_counts.index, y=status_counts.values, ax=axes[0], palette=colors)
axes[0].set_xlabel("Status", fontsize=12, fontweight="bold", color="black")
axes[0].set_ylabel("Count", fontsize=12, fontweight="bold", color="black")
axes[0].set_title("Distribution of Status Values", fontsize=14, fontweight="bold", color="black")
axes[0].tick_params(axis='x', rotation=45, colors="black")
axes[0].tick_params(axis='y', colors="black")
axes[0].grid(False)  # Remove grid from first bar plot

# Add count labels on bars
for bar in bars.patches:
    height = bar.get_height()
    axes[0].text(bar.get_x() + bar.get_width()/2, height, f'{int(height)}', ha='center', va='bottom', fontsize=12, color='black', fontweight="bold")
    bar.set_edgecolor('black')  
    bar.set_linewidth(1.5)  

# Add legend to the first bar chart
legend_handles = [plt.Rectangle((0, 0), 1, 1, color=color) for color in status_colors.values()]
axes[0].legend(legend_handles, status_colors.keys(), title="Status", fontsize=10, title_fontsize=12)

# Pie Chart
explode_values = [0.05] * len(status_counts)
status_counts.plot(
    kind='pie', autopct='%1.1f%%', colors=colors, ax=axes[1],
    wedgeprops={'edgecolor': 'black', 'linewidth': 1.5}, explode=explode_values
)
axes[1].set_title("Status Distribution", fontsize=14, fontweight="bold", color="black")
axes[1].set_ylabel("")  
axes[1].legend(labels=status_counts.index.tolist(), title="Status", loc="upper right", fontsize=10, title_fontsize=12)

# CW Vendor vs. Status (Adding Static Line Between Vendors)
vendor_status_counts = df.groupby(["CW Vendor", "Status"]).size().unstack(fill_value=0)

# Sort vendors by total count
vendor_status_counts = vendor_status_counts.loc[vendor_status_counts.sum(axis=1).sort_values(ascending=False).index]

# Set high RGB colors for vendor plot
sns.set_palette([status_colors.get(status, "gray") for status in vendor_status_counts.columns])  
vendor_status_counts.plot(kind='bar', ax=axes[2], edgecolor="black", linewidth=1.5, width=0.8)

# Labels and titles
axes[2].set_xlabel("CW Vendor", fontsize=12, fontweight="bold", color="black")
axes[2].set_ylabel("Status Count", fontsize=12, fontweight="bold", color="black")
axes[2].set_title("CW Vendor Status", fontsize=14, fontweight="bold", color="black")
axes[2].tick_params(axis='x', rotation=45, colors="black")
axes[2].tick_params(axis='y', colors="black")
axes[2].grid(False)  # Remove grid from second bar plot

# Add count labels on bars
for container in axes[2].containers:
    axes[2].bar_label(container, fmt="%d", fontsize=10, color="black", fontweight="bold")

# Add legend for status categories
axes[2].legend(title="Status", fontsize=10, title_fontsize=12)

# Adding Static Vertical Lines Between CW Vendors
x_positions = np.arange(len(vendor_status_counts)) + 0.5  
for x in x_positions[:-1]:  
    axes[2].axvline(x=x, color='black', linestyle='--', linewidth=1.5)

# Adjust layout for better spacing
plt.tight_layout(rect=[0, 0, 1, 0.96])  # Leave space for title

# Show the final visualization
plt.show()
