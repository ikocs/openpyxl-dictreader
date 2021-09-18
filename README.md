# Excel sheet reader to dictionary
A special class that reads Excel strings and implements access to them through a dictionary: `{Column name: value}`. Implemented on the principle of generators through the `__next__` method.

🇷🇺 [Russian README](README_ru.md) 

Is able to:
* You can define your own fields using the `fieldnames` attribute. If there are more or less own fields, it fills `None` empty spaces, similar to` csv.DictReader () `.
* You can set the string with which to read the column names. In this case, it will start displaying rows after the selected row of column names.
* Fully annotated.

### Implementation example
The example works in conjunction with the file `test.xlsx`.
```python
from dict_reader import XlDictReader
from openpyxl import load_workbook

wb = load_workbook('test.xlsx')
d = XlDictReader(wb['main'])
for r in d:
    v = r["column 1"]
    print(f'{v.value} - {d.line_num}')
```
The result of the script:
```
>>> 1 - 2
>>> None - 3
>>> 12.5 - 4
>>> 2021-02-01 00:00:00 - 5
>>> None - 6
>>> exit - 7
```

### Libraries used
* openpyxl 3.0.8