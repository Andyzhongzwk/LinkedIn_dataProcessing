import pandas as pd

df = pd.read_excel(filePath_in)

print("Before filter...")
num_rows = df.shape[0]
num_col = df.shape[1]
print("Number of rows:", num_rows)
print("Number of columns:", num_col)
print()

# Remove rows with missing values in the 'fullName' or 'linkedinProfileUrl' columns
df = df.dropna(subset=['fullName', 'linkedinProfileUrl'])

# Only keep the people who have studied in polyU
df = df[df['school'].str.contains('Hong Kong Polytechnic University|PolyU|香港理工', na=False, case=False) |
                   df['school2'].str.contains('Hong Kong Polytechnic University|PolyU|香港理工', na=False, case=False) |
                   df['company'].str.contains('Hong Kong Polytechnic University|PolyU|香港理工', na=False, case=False)]

print("After filtered school...")
num_rows = df.shape[0]
print("Number of rows:", num_rows)
print()

# Only keep the people who have studied in Bachelor of Arts related degree
df = df[df['schoolDegree'].str.contains('Arts|English|Linguistics|文學|文学|語言|英文|语言', na=False, case=False) |
                   df['schoolDegree2'].str.contains('Arts|English|Linguistics|文學|文学|語言|英文|语言', na=False, case=False)]

print("After filtered degree...")
num_rows = df.shape[0]
print("Number of rows:", num_rows)
print()

# Export the result
df.to_excel(filePath_out)
