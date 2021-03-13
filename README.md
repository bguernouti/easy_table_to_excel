# Easy Table to xls sheet
```
import random 
s = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
data = {
    "2013": [random.randint(1, 100) for _ in range(0, 8)],
    "2014": [random.randint(1, 100) for _ in range(0, 8)],
    "2015": [random.randint(1, 100) for _ in range(0, 8)],
    "2016": [random.randint(1, 100) for _ in range(0, 8)],
}
index = [
    s[random.randint(0, 20): random.randint(21, 51)] for _ in range(0, 8)
]<br>
df = pd.DataFrame(data=data, index=index)
```
