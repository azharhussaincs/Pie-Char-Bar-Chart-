import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

df = pd.read_excel('NPS CP Data Feb-25.xlsx')

status_array = df["Status"].dropna().values  
status_counts = pd.Series(status_array).value_counts()

plt.figure(figsize=(8, 5))
sns.barplot(x=status_counts.index, y=status_counts.values, hue=status_counts.index, legend=False, palette="viridis")
plt.xlabel("Status")
plt.ylabel("Count")
plt.title("Distribution of Status Values")
plt.xticks(rotation=45)
plt.savefig("status_barplot.png")  
plt.close()

plt.figure(figsize=(6, 6))
status_counts.plot(kind='pie', autopct='%1.1f%%', colors=sns.color_palette("viridis"))
plt.title("Status Distribution")
plt.ylabel("")
plt.savefig("status_piechart.png")  
plt.close()
