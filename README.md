# SciWritingFuncs
Some useful python functions for sci paper writing.

### Dependencies
You need to install some python packages before using these functions:
```
pip install pandas python-docx
```

### Getting started
The quick way to use it is as follows:
```
import pandas as pd
from sci_writing_utils import df_to_three_line_table

df_example = pd.read_csv('/path/to/your/.csv')
df_to_three_line_table(
  df_example,
  output_path='/output/of/your/.docx',
  table_title='example table'
)
```
