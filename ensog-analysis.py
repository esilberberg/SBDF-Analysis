import pandas as pd

ENSOG_file = 'ENSOG Library Catalog Sala de lectura.xlsx'
dewey_codes_file = 'code_constructors/dewey-codes.xlsx'

df = pd.read_excel(ENSOG_file)
dewey_codes_df = pd.read_excel(dewey_codes_file)

def format_dewey_number(call_number):
    """Clean and format a call number to the general class of a Dewey call number"""
    if pd.isna(call_number):
        return None
    
    # Remove alphabetical characters and '-'
    cleaned_number = ''.join(filter(str.isdigit, str(call_number)))
    
    # Extract first 3 digits to isolate Dewey code
    dewey_class = cleaned_number[:3]
    
    # If the cleaned dewey class is 2 digits, add a leading '0'
    if len(dewey_class) == 2:
        dewey_class = '0' + dewey_class

    return dewey_class

def lookup_dewey_class(dewey_number):
    """Look up the description of an associated Dewey class"""
    if dewey_number is not None and dewey_number.strip():  # Check if not None and not an empty string
        try:
            description_series = dewey_codes_df[dewey_codes_df['code'] == int(dewey_number)]['description']
            
            # Extract the first value from the series
            description = description_series.iloc[0] if not description_series.empty else None
            return description
        except ValueError:
            # Handle case where conversion to int fails
            print(f"Invalid format: {dewey_number}")
    
    return None  # Return None if dewey_number is None or an empty string


df['dewey_class'] = df['signatura-topografica'].apply(format_dewey_number)
dewey_class_count = df['dewey_class'].value_counts()

books_by_dewey_class_df = pd.DataFrame({
    'Dewey_code': dewey_class_count.index,
    'Count': dewey_class_count.values,
    'Description': dewey_class_count.index.map(lookup_dewey_class)
})

books_by_author = df['autor'].value_counts()
books_by_publisher = df['editorial'].value_counts()

print(books_by_dewey_class_df)
print(books_by_author)
print(books_by_publisher)

with pd.ExcelWriter('ensog_analysis.xlsx') as writer:
    books_by_dewey_class_df.to_excel(writer, sheet_name='Class', index=False)
    books_by_author.to_excel(writer, sheet_name='Author')
    books_by_publisher.to_excel(writer, sheet_name='Publisher')