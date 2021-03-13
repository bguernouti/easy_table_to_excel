# Easy Table to xls sheet
```
import random 
import pandas as pd
s = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
data = {
    "2013": [random.randint(1, 100) for _ in range(0, 8)],
    "2014": [random.randint(1, 100) for _ in range(0, 8)],
    "2015": [random.randint(1, 100) for _ in range(0, 8)],
    "2016": [random.randint(1, 100) for _ in range(0, 8)],
}
index = [
    s[random.randint(0, 20): random.randint(21, 51)] for _ in range(0, 8)
]
df = pd.DataFrame(data=data, index=index)
```

`print(df)`

```
2013  2014  2015  2016
nopqrstuvwxyz                                  37    47    63    95
ghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVW    48    90    32    70
rstuvwxyzABCDEFGHI                              5    65    63    72
jklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRST          52    46    58    64
bcdefghijklmnopqrstuvwxyzABCDEFGHIJKLM         12    85    18    83
abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLM         8    96    86    48
hijklmnopqrstuvwxyzABCDEFGHIJK                 31    15    41    47
ijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWX      1    87    93    29
```
