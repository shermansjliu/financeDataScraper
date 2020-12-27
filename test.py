import re

table_text = ["estimated values 0.82"]

print([re.findall("\d+\.\d+", t) for t in table_text])