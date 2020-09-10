#######################################################

import pandas as pd
from datetime import datetime
########################################################

path = 'C:\\Users\\User\\Desktop\\project.xlsx'

one_kospi = pd.read_excel(path, sheet_name='코스피', index_col=0)
one_kosdaq = pd.read_excel(path, sheet_name='코스닥', index_col=0)

fwd_op = pd.read_excel(path, sheet_name='영업이익컨센', index_col=0)
op= pd.read_excel(path, sheet_name='영업이익', index_col=0)
op = op.T

fwd_tp = pd.read_excel(path, sheet_name='목표주가괴리율', index_col=0)

#name = pd.read_excel(path, sheet_name='종목', index_col=0)
#name = name.to_dict()
#name = name['종목명']

data = pd.read_excel(path, sheet_name='가격', index_col=0) # 가격데이터

########################################################
#%%
def backtest(op, fwd_op, fwd_tp, data, money, start, rebalancing):
    
    ########################### 매출액 성장률, 목표주가 괴리율 #########################
    back_fwd_op = fwd_op.loc[start] # back_fwd_op = fwd_op.loc[start]
    back_fwd_op = pd.DataFrame(back_fwd_op)
    back_fwd_op.columns = ["컨센영업이익"]
        

    back_op = pd.DataFrame(op)
    back_op.columns = ["영업이익"]

    back_minus = pd.merge(back_op,back_fwd_op,  how='inner', left_index=True, right_index=True)
    
    back_minus['증감률'] = back_minus['컨센영업이익'] / back_minus['영업이익'] -1

    back_minus= back_minus.sort_values(by='증감률', ascending = False)

    back_minus = back_minus[back_minus['증감률'] > 0.40]
    
    back_fwd_tp = fwd_tp.loc[start]
    back_fwd_tp = pd.DataFrame(back_fwd_tp)
    back_fwd_tp.columns = ["목표주가괴리율"]

    back_final = pd.merge(back_minus,back_fwd_tp,  how='inner', left_index=True, right_index=True)
    back_final = back_final[back_final["목표주가괴리율"] > 0.30]
    back_final = back_final.sort_values(by= "목표주가괴리율", ascending = False)
    back_final = back_final.iloc[0:30]
    #################################################

    '''
    백테스트 함수(data = 가격 데이터 csv,
    money = 초기 시작돈, start = 시작날짜
    rebalacning = 리밸런싱 기간(영업일 기준, 한달 = 20으로 !)
    '''
    
    
    etf_price = data[back_final.index]
    etf_price = etf_price.dropna(axis = 1)
    # 받을 가격 data를 etf_price에 옮김.
    # 함수로 계속 리밸런싱을 돌려야하므로! 함수 시작자리에 넣어줌.
    
    mmt_60 = pd.DataFrame(etf_price.pct_change(60).loc[start])
    mmt_60.columns = ['모멘텀']
    mmt_20 = pd.DataFrame(etf_price.pct_change(20).loc[start])
    mmt_20.columns = ['모멘텀']
    
    mmt = mmt_60['모멘텀'] - mmt_20['모멘텀']
    mmt = pd.DataFrame(mmt)
    # 영업일 기준 몇개월 모멘텀?
    # 1개월 모멘텀
    # 일간 수익률을 pct_change로 구하고 인자()는 숫자만큼의 
    # 위의 행 대비 현재 수익률을 나타냄. 20이면, 20 영업일 전의 주가 대비
    # 현재의 수익률을 표시해줌.
        
    mmt['모멘텀순위'] = mmt['모멘텀'].rank(ascending = False) 
    
    # False 일시 + 모멘텀
    # True 일시 - 모멘텀
    
    mmt = mmt.sort_values(by='모멘텀순위')
    # 모멘텀 순위로 정렬하기!
    
    mmt_data = mmt.iloc[0:5]

    mmt_list = mmt_data.index
    
    mmt_price = etf_price[mmt_list][start:]
#####################################    
    
    pf_stock_num = {}
    stock_amount = 0
    

    length = int(len(mmt_list))

    each_money = int(money / length)
    
    for code in mmt_price.columns:
        temp = int(each_money / mmt_price[code][0])
        pf_stock_num[code] = temp
        stock_amount = stock_amount + temp * mmt_price[code][0]

    cash_amount = money - stock_amount


    stock_pf = 0
    
    for code in mmt_price.columns:
        stock_pf = stock_pf + mmt_price[code] * pf_stock_num[code]

    back = pd.DataFrame({'ETF':stock_pf[:rebalancing]})
    

    back['현금'] = [cash_amount] * len(back)

    back['종합'] =back['ETF'] + back['현금']

    back['일별수익률'] = back['종합'].pct_change()
    back['총수익률'] = back['종합']/money - 1

    money = back.iloc[-1,2]
    
    print(date)
    print(mmt_list)
    
    return back, money;

#%%
#################################################
    
start = "2020-04-01"
endday = '2020-08-26'
money = 100000000
backsum = None
plot = 20

#%%
############################################################

for date, i in zip(data[start:endday].index, range(len(data[start:endday].index))):
    

    if i % plot == 0:
        
        back, money = backtest(op, fwd_op, fwd_tp, data, money, date, plot)

        if i == 0 :
            backsum = back
        else:
            backsum = pd.concat([backsum, back])
        
############################################################
            
backsum['일별수익률'] = backsum['종합'].pct_change()
backsum['총수익률'] = backsum['종합'] / backsum['종합'][0] - 1

#############################################################
#%%    
import matplotlib.pyplot as plt

kospi = one_kospi[start:]
kosdaq = one_kosdaq[start:]
kospi['일별수익률'] = kospi['종가지수'].pct_change() 
kospi['총수익률'] = kospi['종가지수'] / kospi['종가지수'][start] - 1

kosdaq['일별수익률'] = kosdaq['종가지수'].pct_change() 
kosdaq['총수익률'] = kosdaq['종가지수'] / kosdaq['종가지수'][start] - 1

plt.figure(figsize=(10,6))

kospi['총수익률'].plot(label='KOSPI')
kosdaq['총수익률'].plot(label='KOSDAQ')
backsum['총수익률'].plot(label='Untact ETF')
plt.legend()
plt.show()
