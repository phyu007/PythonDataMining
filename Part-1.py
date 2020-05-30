# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
# read the docx file
# Locate the value
# found ? store them separately
# loop (append txt file with | each )


# %%
from docx import Document
from docx.api import Document
import pandas as pd
import docx2txt

# extract text
text = docx2txt.process(
    r"./Sample term sheet 2016.docx")

print(text)


# %%
content = []
for line in text.splitlines():
    if line != '':
        content.append(line)

print(content)
print(content[0])


# %%
keyword_list = ['Trade Date:', 'Initial Valuation Date:',
                'Effective Date:', 'Notional Amount:', 'Fixed Rate:']
answer = []


for con in content:
    if "TERM SHEET" in con:
        print(content.index(con))
        answer.append(content[content.index(con)].split()[0])

for con in content:
    if "Bank ref:" in con:
        print(content.index(con))
        answer.append("Bank"+content[content.index(con)].strip("Bank ref:"))

for con in content:
    if con in keyword_list:
        print(content.index(con))
        answer.append(content[content.index(con)+1])

print(answer)


# %%
with open(r"./key_attributes.txt", "w") as txt_file:
    for line in answer:
        txt_file.write(" ".join(line) + "|")


# %%
# b

# get the table value
# construct dataframe
# save it to the text file


# %%

document = Document(
    r"./Sample term sheet 2016.docx")


table = document.tables[5]


data = []

keys = None
for i, row in enumerate(table.rows):
    text = (cell.text for cell in row.cells)

    if i == 0:
        keys = tuple(text)
        continue
    row_data = dict(zip(keys, text))
    data.append(row_data)

print(data)
df = pd.DataFrame(data)
print(df)


# %%
a = df.iloc[:-1, 1:]
ans = a.to_numpy()
print(ans)


# %%
with open(r"./payment_date.txt", "w") as txt_file:
    for row in ans:
        txt_file.write("|".join(row) + "\n")


# %%

