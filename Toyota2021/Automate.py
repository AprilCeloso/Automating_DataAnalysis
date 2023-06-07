# %%
import pandas as pd

# %%
file_path = input("* Enter here the file location of you file : ")


# %%
df = pd.read_excel(file_path)

# %%
unique_sort_values = df['Sort'].unique()

# %%
results_df = pd.DataFrame(columns=['Car Company', 'Total Headlines', 'Total Ad Value'])

# %%
for sort_value in unique_sort_values:
    filtered_data = df[df['Sort'] == sort_value]
    total_headlines = len(filtered_data)
    total_ad_value = filtered_data['Ad Value'].sum()
    
    results_df = results_df.append({'Car Company': sort_value, 'Total Headlines': total_headlines, 'Total Ad Value': total_ad_value}, ignore_index=True)

# %% [markdown]
# 

# %%
results_df = results_df.sort_values(by='Total Headlines', ascending=False)

# %%
top_3_rows = results_df.head(3)

# %%
top_3_rows

# %% [markdown]
# TASK 3

# %%
df['MV Name'] = df['MV Name'].str.strip().replace('', 'No Author')

# %%
toyota_data = df[df['Sort'] == 'Toyota']

# %%
toyota_summary = toyota_data.groupby('MV Name')['AVE Value'].sum().reset_index()


# %%
sorted_mv_names = toyota_summary.sort_values('AVE Value', ascending=False)


# %%
names= sorted_mv_names.head(3)

# %%
names


# %% [markdown]
# TASK 2

# %%
toyota_data = df[df['Sort'] == 'Toyota']

# %%
headlines_summary = toyota_data.groupby('Headline')['AVE Value'].sum().reset_index()
sorted_headlines = headlines_summary.sort_values('AVE Value', ascending=False)


# %%
top_3_headlines = sorted_headlines.head(3)
print("Top 3 Headlines for Toyota:")
print(top_3_headlines[['Headline', 'AVE Value']].reset_index(drop=True))

# %%
non_toyota_data = df[df['Sort'] != 'Toyota']

# %%
company_summary = non_toyota_data.groupby('Sort')['AVE Value'].sum().reset_index()
sorted_companies = company_summary.sort_values('AVE Value', ascending=False)


# %%
top_3_companies = sorted_companies.head(3)

# %%
top_3_headlines = non_toyota_data[non_toyota_data['Sort'].isin(top_3_companies['Sort'])].groupby('Sort')['Headline'].first().reset_index()
top_3_companies_with_headlines = pd.merge(top_3_companies, top_3_headlines, on='Sort')

# %%
print("Top 3 Companies (excluding Toyota) and their Top Headlines:")
print(top_3_companies_with_headlines[['Sort', 'Headline', 'AVE Value']].reset_index(drop=True))

# %% [markdown]
# 

# %%
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    writer.book = writer.book
    top_3_headlines.to_excel(writer, sheet_name='Task 2', index=False)
    top_3_companies_with_headlines.to_excel(writer, sheet_name='Task 2.1', index=False)
    top_3_rows.to_excel(writer, sheet_name='Task 1', index=False)
    names.to_excel(writer, sheet_name='Task 3', index=False)


