#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd
import numpy as np


# In[8]:


print(np.arange(9).reshape(3, 3))


# In[32]:


dataframe = pd.DataFrame(np.arange(30).reshape(10, 3),columns=['商品名', '価格', '個数'])
print(dataframe)


# In[ ]:


menu_array = ['DXハンバーグ弁当', 'ハンバーグ弁当', '春巻き&ササミカツ弁当', 'ガパオライス', 'オムライス', 'サバカラ弁当', 'シャケカラ弁当', 'カラアゲ弁当']
print(dataframe)
for index, product_name in enumerate(menu_array):
    dataframe[index]['商品名'] = product_name
print(dataframe)

