import pandas as pd
from collections import Counter

# Load the Excel file into a DataFrame
file_path = 'Well-Being Survey Offboarding(1-297).xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Columns for frequency analysis
col_most_helpful = 'Please select three components of the Longevity Program that you helped you the most to achieve your health and wellness goals\n'
col_least_helpful = 'Please select three components of the Longevity Program that you helped you the least to achieve your health and wellness goals\n'
col_onboarding_experience = 'On a scale of 1-10, how comfortable was the onboarding experience (1 being lowest and 10 being highest)\n'
col_recommend_program = 'How likely are you to recommend the Longevity Program to one of your friends or family members?\n'
col_agree_statement = 'On a scale of 1-10 (1= Completely Disagree, 10=Completely Agree), please indicate how much you agree with the following statement, “The Longevity trial has given me the information and motivation ...'

# Frequency counts for most and least helpful components
freq_most_helpful = df[col_most_helpful].value_counts().reset_index().rename(columns={'index': 'Most Helpful Components', col_most_helpful: 'Frequency'})
freq_least_helpful = df[col_least_helpful].value_counts().reset_index().rename(columns={'index': 'Least Helpful Components', col_least_helpful: 'Frequency'})

# Bin edges and labels for the categorization of scaled questions
bin_edges = [0, 2, 4, 6, 8, 10]
bin_labels_experience = ['Very Poor', 'Poor', 'Average', 'Very Good', 'Excellent']
bin_labels_recommend = ['Highly Unlikely', 'Unlikely', 'Neutral', 'Likely', 'Highly Likely']
bin_labels_agree = ['Strongly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']

# Create new columns with categorized responses
df['Onboarding Experience Category'] = pd.cut(df[col_onboarding_experience], bins=bin_edges, labels=bin_labels_experience, right=True)
df['Likelihood to Recommend Category'] = pd.cut(df[col_recommend_program], bins=bin_edges, labels=bin_labels_recommend, right=True)
df['Agreement with Statement Category'] = pd.cut(df[col_agree_statement], bins=bin_edges, labels=bin_labels_agree, right=True)

# Generate frequency distributions for these new categorized columns
freq_onboarding_experience_cat = df['Onboarding Experience Category'].value_counts().reset_index().rename(columns={'index': 'Onboarding Experience Category', 'Onboarding Experience Category': 'Frequency'})
freq_recommend_program_cat = df['Likelihood to Recommend Category'].value_counts().reset_index().rename(columns={'index': 'Likelihood to Recommend Category', 'Likelihood to Recommend Category': 'Frequency'})
freq_agree_statement_cat = df['Agreement with Statement Category'].value_counts().reset_index().rename(columns={'index': 'Agreement with Statement Category', 'Agreement with Statement Category': 'Frequency'})

# Split the string responses into lists and flatten them for most and least helpful components
most_helpful_list = df[col_most_helpful].str.split(';').explode().dropna()
least_helpful_list = df[col_least_helpful].str.split(';').explode().dropna()

# Count the frequencies using Counter
most_helpful_count = Counter(most_helpful_list)
least_helpful_count = Counter(least_helpful_list)

# Convert the Counter objects to DataFrames for easier visualization
df_most_helpful_split = pd.DataFrame.from_dict(most_helpful_count, orient='index', columns=['Frequency']).reset_index().rename(columns={'index': 'Most Helpful Components'})
df_least_helpful_split = pd.DataFrame.from_dict(least_helpful_count, orient='index', columns=['Frequency']).reset_index().rename(columns={'index': 'Least Helpful Components'})

# Initialize Excel writer object
output_file_path = 'survey analysis 05102023'  # Replace with your desired output file path
writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

# Write each DataFrame to specific sheets in the Excel file
freq_most_helpful.to_excel(writer, sheet_name='Most Helpful Components', index=False)
freq_least_helpful.to_excel(writer, sheet_name='Least Helpful Components', index=False)
freq_onboarding_experience_cat.to_excel(writer, sheet_name='Onboarding Experience Categories', index=False)
freq_recommend_program_cat.to_excel(writer, sheet_name='Likelihood to Recommend Categories', index=False)
freq_agree_statement_cat.to_excel(writer, sheet_name='Agreement with Statement Categories', index=False)
df_most_helpful_split.to_excel(writer, sheet_name='Most Helpful Split', index=False)
df_least_helpful_split.to_excel(writer, sheet_name='Least Helpful Split', index=False)

# Save the Excel file
writer.save()