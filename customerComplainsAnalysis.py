import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# Load and clean the data
file_path = 'customerComplainsData.xlsx'  # Update this with the actual path to your Excel file
data = pd.read_excel(file_path)
data['Date submitted'] = pd.to_datetime(data['Date submitted'])
data['Date received'] = pd.to_datetime(data['Date received'])
data.fillna('Unknown', inplace=True)
data.drop_duplicates(inplace=True)

# Seasonal Patterns Q1
# Assuming data is already loaded and cleaned

# Extract year and month from 'Date submitted'
data['Year'] = data['Date submitted'].dt.year
data['Month'] = data['Date submitted'].dt.month

# Group by year and month, then count the complaints
complaints_by_month_year = data.groupby(['Year', 'Month']).size().reset_index(name='Complaints')

# Create a pivot table for better visualization
pivot_table = complaints_by_month_year.pivot(index="Month", columns="Year", values="Complaints")

# Plot the data
plt.figure(figsize=(12, 8))
sns.heatmap(pivot_table, cmap="YlGnBu", annot=True, fmt="g")
plt.title('Monthly # Of Complaints Over the Years')
plt.xlabel('Year')
plt.ylabel('Month')
plt.yticks(ticks=range(12), labels=['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], rotation=0)
plt.show()

# Assuming the heatmap and pivot_table have already been generated

# Overall Trend Summary
print("Overall Trend Summary:")
total_complaints_by_year = complaints_by_month_year.groupby('Year')['Complaints'].sum()
print(f"Total complaints by year:\n{total_complaints_by_year}\n")

average_monthly_complaints = complaints_by_month_year.groupby('Month')['Complaints'].mean()
print(f"Average complaints by month:\n{average_monthly_complaints}\n")

# Highlight Specific Observations
max_complaints_month_year = complaints_by_month_year.loc[complaints_by_month_year['Complaints'].idxmax()]
print(f"Month with highest complaints: {max_complaints_month_year['Month']} in {max_complaints_month_year['Year']}, with {max_complaints_month_year['Complaints']} complaints.\n")

min_complaints_month_year = complaints_by_month_year.loc[complaints_by_month_year['Complaints'].idxmin()]
print(f"Month with lowest complaints: {min_complaints_month_year['Month']} in {min_complaints_month_year['Year']}, with {min_complaints_month_year['Complaints']} complaints.\n")

# Year-over-Year Comparison for a Specific Month (e.g., January)
jan_complaints = complaints_by_month_year[complaints_by_month_year['Month'] == 1]
print("Year-over-Year Comparison for January:")
print(jan_complaints.set_index('Year')['Complaints'])



# Most Complained Products and Issues Q2
product_complaints = data['Product'].value_counts().head(10)
top_product = product_complaints.index[0]
top_product_issues = data[data['Product'] == top_product]['Issue'].value_counts().head(10)

fig, ax = plt.subplots(1, 2, figsize=(18, 6))

# Top 10 products with the most complaints
sns.barplot(x=product_complaints.values, y=product_complaints.index, palette='coolwarm', ax=ax[0])
ax[0].set_title(f'Top 10 Products with Most Complaints (2017-2023)')

# Add data labels for the first barplot
for i, v in enumerate(product_complaints.values):
    ax[0].text(v + 3, i, str(v), color='black', va='center')

# Top 10 issues for the product with the most complaints
sns.barplot(x=top_product_issues.values, y=top_product_issues.index, palette='viridis', ax=ax[1])
ax[1].set_title(f'Top 10 Issues for "{top_product}"')

# Add data labels for the second barplot
for i, v in enumerate(top_product_issues.values):
    ax[1].text(v + 3, i, str(v), color='black', va='center')

plt.tight_layout()
plt.show()

# Complaint Resolution Q3
resolution_types = data['Company response to consumer'].value_counts()
plt.figure(figsize=(10, 6))
resolution_plot = sns.barplot(x=resolution_types.values, y=resolution_types.index, palette='Set2')
plt.title('Company Responses to Consumer Complaints (2017-2023)')

# Add data labels for the resolution_types barplot
for i, v in enumerate(resolution_types.values):
    plt.text(v + 3, i, str(v), color='black', va='center')

plt.show()

# Untimely Responses Q4
untimely_complaints = data[data['Timely response?'] == 'No']
untimely_by_product = untimely_complaints['Product'].value_counts().head(10)
plt.figure(figsize=(10, 6))
untimely_plot = sns.barplot(x=untimely_by_product.values, y=untimely_by_product.index, palette='magma')
plt.title('Top 10 Products with Untimely Complaint Responses (2017-2023)')

# Add data labels for the untimely_by_product barplot
for i, v in enumerate(untimely_by_product.values):
    plt.text(v + 3, i, str(v), color='black', va='center')

plt.show()


# Sample data
categories = ['A', 'B', 'C', 'D']
values = [10, 20, 15, 5]

# Get a color from the 'coolwarm' colormap
color = plt.cm.coolwarm(0.5)  # Extract the RGBA color

# Create a simple bar plot
plt.figure(figsize=(8, 4))
sns.barplot(x=categories, y=values, palette=[color])

# Show the plot
plt.show()