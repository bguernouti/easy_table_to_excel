# Easy Table to xls sheet
```
import random 
import pandas as pd
index = ["Net income", "Minority rights","Net revenue", "Net loans", "Total debt", "Earnings", "consumption", "Treasury stocks"]
random.shuffle(index)
data = {
    "2013": [random.randint(1, 100) for _ in range(0, 8)],
    "2014": [random.randint(1, 100) for _ in range(0, 8)],
    "2015": [random.randint(1, 100) for _ in range(0, 8)],
    "2016": [random.randint(1, 100) for _ in range(0, 8)],
}
df = pd.DataFrame(data=data, index=index)
```
`print(df)`
```
2013  2014  2015  2016
Net income        100    91    87    86
Net revenue        99    87    22    54
Net loans          10    51    35    93
consumption        41    46    12    70
Treasury stocks    52     8    11    48
Earnings           61    64    98     8
Total debt         31    74    37   100
Minority rights    36    77    79    98
```
