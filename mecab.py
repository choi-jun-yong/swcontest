import pandas as pd
from collections import Counter
import re
from PyKomoran import *
from konlpy.tag import *
from konlpy.tag import Mecab
import pandas as pd
import os




#f = open('C:\\Users\\Desktop\\텍스트데이터.txt', "r")
#lines = f.read()

list_company = pd.read_excel('C:\\Users\\User\\Desktop\\기업명.xlsx', index_col = 0)
list_text = pd.read_excel('C:\\Users\\User\\Desktop\\abc.xlsx', index_col = 0)

m = Mecab()

dict_word = {}

for i in range(len(list_text)):
    sentence = list_text.iloc[i,0]
    temp = m.nouns(sentence)
    count = Counter(temp)
    dict_temp = list(dict(count).keys())
    for j in range(len(dict_temp)):
        temp_count_word = dict_word.get(dict_temp[j],0) # 없으면 0 있으면 value
        temp_count_word += 1
        dict_word[dict_temp[j]] = temp_count_word

wordlist = dict_word.keys()

################################
final_company_list = []

for i in range(len(list_company.index)):
    final_company_list.append(list_company.iloc[i,0])
##################################
    
def isCorp(noun):
    return noun in final_company_list

company_word_one = list(filter(isCorp, wordlist))

######################################

final_companylist = {}

for i in range(len(company_word_one)):
    final_companylist[company_word_one[i]] = dict_word[company_word_one[i]]

final_companylist = pd.DataFrame.from_dict(final_companylist, orient = 'index')
final_companylist.columns = ['빈도수']
final_companylist = final_companylist.sort_values(by = '빈도수', ascending=False)

os.getcwd()
os.chdir('C:\\Users\\최준용\\Desktop')
final_companylist.to_excel("test.xlsx")
