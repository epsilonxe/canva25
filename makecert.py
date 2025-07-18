import pandas as pd
import numpy as np

df = pd.read_excel('canva25.xlsx')
name = df['ชื่อ'].str.cat(df['นามสกุล'], sep=' ').to_numpy()
name = np.strings.replace(name, u'\xa0', u'')
print(name)