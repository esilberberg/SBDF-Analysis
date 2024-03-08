import re
import pandas as pd

def parse_text(text):
    # Regular expression finds segments with a three-digit code followed by a description
    pattern = re.compile(r'(\d{3})\s(.+?)(?=\d{3}\s|$)')
    matches = re.findall(pattern, text)

    # List of dictionaries with code and description
    result = [{'code': match[0], 'description': match[1].strip()} for match in matches]

    return result

input_text = 'dewey.txt'
with open(input_text, 'r') as file:
        file_contents = file.read()


result_list = parse_text(file_contents)

df = pd.DataFrame(result_list)

df.to_excel('output.xlsx', index=False)