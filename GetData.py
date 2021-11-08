import tushare as ts
import pandas
import openpyxl
import pymysql

from pymysql.converters import escape_string

#初始化接口
ts.set_token('fec60fbea9722d575c28a73e0a0b85cc9d3712a8f79a4f871832c317')

db =pymysql.connect(host="localhost",user="root",password="950910",database="stock")
cursor =db.cursor()
pro = ts.pro_api()
wb = openpyxl.Workbook()
sheet = wb.active
sheet.cell(row=1,column=1,value = '代码')
sheet.cell(row=1,column=2,value = '买入日期')
sheet.cell(row=1,column=3,value = '收盘价')
sheet.cell(row=1,column=4,value = '成交量')
sheet.cell(row=1,column=5,value = '20均价')
sheet.cell(row=1,column=6,value = 'C/MA20')
sheet.cell(row=1,column=7,value = '卖出日期')
sheet.cell(row=1,column=8,value = '收盘价')
sheet.cell(row=1,column=9,value = '成交量')
sheet.cell(row=1,column=9,value = '20均价')
sheet.cell(row=1,column=10,value = '差价')
sheet.cell(row=1,column=11,value = '盈亏比')
x = 2
df_code = pro.stock_basic(exchange='',list_status='L',fields='ts_code,name,area,industry,list_date')  #获取股票代码和行业、地域
y_t = 0
y_k = 0

for i in range(int(df_code.size/5)-1):
    #print(df_code)
    #print(int(df_code.size/5))
    #df_daily = pro.daily(ts_code=df_code.iloc[i,0],strat_date='20180101',end_date='20201231')
    df_data = ts.pro_bar(ts_code=df_code.iloc[i,0],adj='qfq',start_date='20180101',end_date='20181231',ma=[20])

    print(df_code.iloc[i,0],df_code.iloc[i,4])
    if (df_code.iloc[i,4]<'20180101') and (int(df_data.size/13)-1>0):
        for j in range(int(df_data.size/13)-1,0,-1):

            if (df_data.iloc[j,11]!=0) and (df_data.iloc[j,5]>=df_data.iloc[j,11]*1.1) and (df_data.iloc[j,5]!=0) and (df_data.iloc[j,5]<=df_data.iloc[j,11]*1.4) and (df_data.iloc[j,5]>df_data.iloc[j,2]):
                y = 0
                for y in range(int(df_data.size/13)-j):
                    if (df_data.iloc[j-y,5]<=df_data.iloc[j,5]*0.95) or (df_data.iloc[j-y,5]<=df_data.iloc[j-y,11]) or (df_data.iloc[j-y,5]>=df_data.iloc[j-y,11]*1.4) :
                        if (df_data.iloc[j - y, 5] <= df_data.iloc[j, 5] * 0.95) or (df_data.iloc[j - y, 5] <= df_data.iloc[j - y, 11]):
                            y_k += 1
                        if  df_data.iloc[j - y, 5] >=df_data.iloc[j - y, 11] * 1.4:
                            y_t += 1
                        print(df_data.iloc[j,1],df_data.iloc[j,5],df_data.iloc[j,11],df_data.iloc[j,5]/df_data.iloc[j,11])
                        sheet.cell(row=x,column=1,value = df_code.iloc[i,0])    #代码
                        sheet.cell(row=x,column=2,value = df_data.iloc[j,1])   #买入交易日期
                        sheet.cell(row=x,column=3,value = df_data.iloc[j,5])   #收盘价
                        sheet.cell(row=x,column=4,value = df_data.iloc[j,10])  #成交量
                        sheet.cell(row=x,column=5,value = df_data.iloc[j,11])    #20均价
                        sheet.cell(row=x,column=6,value = df_data.iloc[j,5]/df_data.iloc[j, 11])  # 20均价
                        sheet.cell(row=x,column=7,value = df_data.iloc[j-y,1])    #卖出日期
                        sheet.cell(row=x,column=8,value = df_data.iloc[j-y,5])    #收盘价
                        sheet.cell(row=x,column=9,value = df_data.iloc[j-y,10])    #成交量
                        sheet.cell(row=x,column=9,value = df_data.iloc[j-y,11])    #20均价
                        sheet.cell(row=x,column=10,value = df_data.iloc[j-y,5]-df_data.iloc[j,5])    #差价
                        sheet.cell(row=x,column=11,value = (df_data.iloc[j-y,5]-df_data.iloc[j,5])/df_data.iloc[j,5])    #盈亏比
                        x +=  1
                        j -= y-1
                        break
#                    if  df_data.iloc[j+y,5]>=df_data.iloc[j+y,11]*1.4:
#                        print(df_data.iloc[j,1],df_data.iloc[j,5],df_data.iloc[j,11])
#                        sheet.cell(row=x,column=1,value = df_code.iloc[i,0])    #代码
#                        sheet.cell(row=x,column=2,value = df_data.iloc[j,1])   #买入交易日期
#                        sheet.cell(row=x,column=3,value = df_data.iloc[j,5])   #收盘价
#                        sheet.cell(row=x,column=4,value = df_data.iloc[j,10])  #成交量
#                        sheet.cell(row=x,column=5,value = df_data.iloc[j,11])    #20均价
#                        sheet.cell(row=x,column=6,value = df_data.iloc[j,5]/df_data.iloc[j, 11])  # 20均价
#                        sheet.cell(row=x,column=7,value = df_data.iloc[j+y,1])    #卖出日期
#                        sheet.cell(row=x,column=8,value = df_data.iloc[j+y,5])    #收盘价
#                        sheet.cell(row=x,column=9,value = df_data.iloc[j+y,10])    #成交量
#                        sheet.cell(row=x,column=9,value = df_data.iloc[j+y,11])    #20均价
#                        sheet.cell(row=x,column=10,value = df_data.iloc[j+y,5]-df_data.iloc[j,5])    #差价
#                        sheet.cell(row=x,column=11,value = df_data.iloc[j,5]/(df_data.iloc[j+y,11]-df_data.iloc[j,5]))    #盈亏比
#                        x = x + 1

#                        break



wb.save('ma20.xlsx')
print('盈利笔数：'+str(y_t))
print('亏损笔数：'+str(y_k))
#print('盈亏比：'+str(y_t/(y_t+y_k)))


#df_code = pro.stock_basic(exchange='',list_strtus='L',fields='ts_code,name,area,industry')
#for i in range(int(df_code.size/4)-1):
    
#    insert_sql = "insert into data-code values ('"+df_code.iloc[i,0]+"','"+df_code.iloc[i,1]+"','"+df_code.iloc[i,2]+"','"+df_code.iloc[i,3]+"')"
                 #%(str(df_code.iloc[i,0]),str(df_code.iloc[i,1]),str(df_code.iloc[i,2]),str(df_code.iloc[i,3]))
#    cursor.execute(insert_sql)


#df_data = ts.pro_bar(ts_code='600887.SH', asset='E', start_date='20180101', end_date='20181011',ma=[20])
#print(df)
#    print(df.iloc[i,0])    # 股票代码

#for i in range(int(df.size/2)):
#    dt=pro.daily(ts_code=df.iloc[i,0],strat_date='20211018',end_date='20211018')
#    dt1.append(dt,ignore_index=True)
#    dt.to_excel('test.xlsx')
#print(dt)

#dt=pro.daily(ts_code=df.iloc[0,0],strat_date='20211001',end_date='20211018')
#for j in range(int(dt.size/11)):
#    if dt.iloc[j,8]<9.98:
#        dt.drop(index=j)
#        print(dt)


#dt.to_excel('test.xlsx')