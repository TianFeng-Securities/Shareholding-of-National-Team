#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os

# In[2]:


print('即将运行国家队数据自动化处理程序\n')
key = input('是否继续？请输入y/Y继续，输入n/N退出\n')
if key == 'n' or key == 'N':
    exit()

path = '.\\result'
for i in os.listdir(path):
	path_file = os.path.join(path,i)
	if os.path.isfile(path_file):
		os.remove(path_file)
# In[3]:


print('请阅读以下内容,确保提供正确的数据源,请注意您提供的数据一定要与模板数据一致,否则程序将会出错！\n')


# In[4]:


input('1.您需要更新template_all_company.xlsx中的全部A股数据，我已经使用wind函数写好，您只需要打开一次，修改日期值为当前季度末然后待数据更新完毕后保存即可，确认后按任意键继续! \n ')
input('2.您需要提供上季度的综合数据，可从上个季度数据复制，格式与template_last.xlsx一致，并将其命名为last.xlsx,请注意后缀一定为.xlsx！确认后按任意键继续!\n')
input('3.您需要提供当前季度的数据，从万得数据库复制，复制后的格式应与template_data.xlsx文件一致，包括证金，汇金，以及中证金融资管的十个基金，命名为data.xlsx，确认后请按任意键继续!\n')
input('确保以上数据准备完毕，可以开始自动化处理数据，请按任意键继续!\n')


# In[5]:


label = input('请输入您需要的季度，如19Q3\n')


# In[6]:


####计算上个季度
if label[-1] == '1':
    last_label = str(int(label[0:2])-1) + 'Q4'
else:
    last_label = label[0:3] + str(int(label[-1])-1)


# In[7]:


try:
    all_company = pd.read_excel('template/template_all_company.xlsx',usecols=['代码','证券简称','行业','板块','期末自由流通市值'])
except:
    print('template_all_company.xlsx数据有错误\n')
    exit()


# In[8]:


try:
    证金 = pd.read_excel('data.xlsx',sheet_name='证金',usecols=['证券简称','期末参考市值(亿元)'])
    汇金 = pd.read_excel('data.xlsx',sheet_name='汇金',usecols=['证券简称','期末参考市值(亿元)'])
    中欧基金 = pd.read_excel('data.xlsx',sheet_name='中欧基金',usecols=['证券简称','期末参考市值(亿元)'])
    银华基金 = pd.read_excel('data.xlsx',sheet_name='银华基金',usecols=['证券简称','期末参考市值(亿元)'])
    易方达基金 = pd.read_excel('data.xlsx',sheet_name='易方达基金',usecols=['证券简称','期末参考市值(亿元)'])
    南方基金 = pd.read_excel('data.xlsx',sheet_name='南方基金',usecols=['证券简称','期末参考市值(亿元)'])
    嘉实基金 = pd.read_excel('data.xlsx',sheet_name='嘉实基金',usecols=['证券简称','期末参考市值(亿元)'])
    华夏基金 = pd.read_excel('data.xlsx',sheet_name='华夏基金',usecols=['证券简称','期末参考市值(亿元)'])
    大成基金 = pd.read_excel('data.xlsx',sheet_name='大成基金',usecols=['证券简称','期末参考市值(亿元)'])
    广发基金 = pd.read_excel('data.xlsx',sheet_name='广发基金',usecols=['证券简称','期末参考市值(亿元)'])
    博时基金 = pd.read_excel('data.xlsx',sheet_name='博时基金',usecols=['证券简称','期末参考市值(亿元)'])
    工银瑞信 = pd.read_excel('data.xlsx',sheet_name='工银瑞信',usecols=['证券简称','期末参考市值(亿元)'])
    梧桐树 = pd.read_excel('data.xlsx',sheet_name='梧桐树',usecols=['证券简称','期末参考市值(亿元)'])
except:
    print('data.xlsx数据有错误，请检查各个sheet_name,文件命名是否有问题\n')
    exit()


# In[9]:


try:
    last = pd.read_excel('last.xlsx',usecols=['代码','证券简称','行业','板块','期末参考市值(亿元)','期末自由流通市值'])
except:
    print('last.xlsx数据有错误，请检查各个列名,文件命名等是否有问题\n')
    exit()


# In[10]:


print('数据读取完毕\n')


# In[11]:


####第一步中证金融十个基金国家队数据计算
十支基金 = ['中欧基金','银华基金','易方达基金','南方基金','嘉实基金','华夏基金','大成基金','广发基金','博时基金','工银瑞信']
def get_中证金融():
    df = all_company.copy()
    df.drop(columns=['代码','期末自由流通市值','行业','板块'],inplace=True)
    df['期末参考市值(亿元)'] = pd.Series([0 for i in range(df.shape[0])])
    for name in 十支基金:####理解起来很难，后面维护的同学注意可以试试一步一步调试，尤其是eval（name）的使用
        df = df.merge(eval(name),on='证券简称',how='left',suffixes=('','_1'))
        df.fillna(0,inplace=True)
        df['期末参考市值(亿元)'] = df['期末参考市值(亿元)']+df['期末参考市值(亿元)_1']
        df.drop(columns=['期末参考市值(亿元)_1'],inplace=True)
    return df
中证金融 = get_中证金融()
###至此，已经生成：
#证金，汇金，中证金融，梧桐树四个国家队基金，columns=['证券简称'，'期末参考市值(亿元)']


# In[12]:


####输出：国家队持股规模
四个国家队 = ['证金','汇金','中证金融','梧桐树']
国家队持股规模 = pd.DataFrame({'持有市值':四个国家队,
                        label:[证金['期末参考市值(亿元)'].sum(),
                               汇金['期末参考市值(亿元)'].sum(),
                               中证金融['期末参考市值(亿元)'].sum(),
                               梧桐树['期末参考市值(亿元)'].sum()]
                       })
国家队持股规模.to_excel('./result/国家队持股规模.xlsx',sheet_name='国家队持股规模',index=False)
print('已经输出国家队持股规模至"国家队持股规模.xlsx"\n')


# In[13]:


####第二步获取四个国家队相加总和数据
def step2():
    df = all_company.copy(deep=True)
    df['期末参考市值(亿元)'] = pd.Series([0 for i in range(df.shape[0])])
    for name in 四个国家队:
        df = df.merge(eval(name),on=['证券简称'],how='left',suffixes=('','_1'))
        df.fillna(0,inplace=True)
        df['期末参考市值(亿元)'] = df['期末参考市值(亿元)']+df['期末参考市值(亿元)_1']
        df.drop(columns=['期末参考市值(亿元)_1'],inplace=True)
    df = df.merge(last[['证券简称','期末参考市值(亿元)','期末自由流通市值']],on=['证券简称'],how='left',suffixes=('','_1'))
    df.rename(columns={'期末参考市值(亿元)_1':last_label+"参考市值",'期末自由流通市值_1':last_label+"自由流通市值"},inplace=True)
    ###处理新进的df.loc[pd.isna(a['19Q4参考市值']),'期末参考市值(亿元)'] !=0，说明存在新进的股票,直接drop_na()是不合适的
    temp = df.loc[pd.isna(df[last_label+'参考市值']),'期末参考市值(亿元)']
    for i in temp.index:
        if temp[i] != 0.0:
            df.loc[i,'期末参考市值(亿元)'] = 0.0
    df.dropna(inplace=True)
    df.reset_index(inplace=True)
    df.drop(columns=['index'],inplace=True)
    df['市值增减'] = df['期末参考市值(亿元)']-df[last_label+'参考市值']
    df['持有市值占比'] = df['期末参考市值(亿元)']/df['期末自由流通市值']
    df[last_label+'持有市值占比'] = df[last_label+'参考市值']/df[last_label+"自由流通市值"]
    return df


# In[14]:


####输出：合并后的当前季度数据
合并后的当前季度数据 = step2()
合并后的当前季度数据.to_excel('./result/'+label+'.xlsx',sheet_name=label,index=False)
print('已经输出'+label+'至'+label+'.xlsx\n')


# In[15]:


####第三步获得国家队每季度个股增减持前十位
def step3():
    增持 = 合并后的当前季度数据.sort_values(by='市值增减',axis=0,ascending=False)
    增持.reset_index(inplace=True)
    增持.drop(columns=['index'],inplace=True)
    增持 = 增持.loc[:9,['代码','证券简称','行业','市值增减']]
    增持.to_excel('./result/个股增持前十.xlsx',index=False)
    print('已经输出个股增持前十数据至个股增持前十.xlsx')
####第四步获得国家队每季度个股增减持前十位
def step4():
    减持 = 合并后的当前季度数据.sort_values(by='市值增减',axis=0,ascending=True)
    减持.reset_index(inplace=True)
    减持.drop(columns=['index'],inplace=True)
    减持 = 减持.loc[:9,['代码','证券简称','行业','市值增减']]
    减持.to_excel('./result/个股减持前十.xlsx',index=False)
    print('已经输出个股坚持前十数据至个股减持前十.xlsx\n')
step3()
step4()


# In[16]:


####第五步，第六步获得每季度增减幅行业前十位
def step5():
    df = 合并后的当前季度数据.groupby(by='行业')['期末自由流通市值','期末参考市值(亿元)',last_label+'参考市值',last_label+"自由流通市值"].sum()
    df['市值增减'] = df['期末参考市值(亿元)']-df[last_label+'参考市值']
    df['持有市值占比'] = df['期末参考市值(亿元)']/df['期末自由流通市值']
    df[last_label+'持有市值占比'] = df[last_label+'参考市值']/df[last_label+"自由流通市值"]
    df['行业'] = df.index
    df['占比增减'] = df['持有市值占比'] - df[last_label+'持有市值占比']
    df.to_excel('./result/行业合并数据.xlsx',index=False)
    print('已经输出行业数据至行业合并数据.xlsx\n')
    return df

行业合并数据 = step5()

def step6(ascending=True):####默认减持
    df = 行业合并数据.copy()
    df = df.sort_values(by='占比增减',ascending=ascending)
    df.drop(columns=['行业'],inplace=True)
    df.reset_index(inplace=True)
    df = df.loc[:9,['行业','市值增减','占比增减']]
    if ascending:
        df.to_excel('./result/行业减持前十.xlsx',index=False)
        print('已经输出行业减持前十数据至行业减持前十.xlsx\n')
    else:
        df.to_excel('./result/行业增持前十.xlsx',index=False)
        print('已经输出行业增持前十数据至行业减持前十.xlsx\n')
step6()
step6(False)
    


# In[17]:


####第七步获得板块分布情况
df_all = all_company.groupby(by='板块')['期末自由流通市值'].sum()

df_持有 = 合并后的当前季度数据.groupby(by='板块')['期末参考市值(亿元)'].sum()
df_持有.to_excel("./result/板块分布_持股市值.xlsx")
print('已经输出持股市值至板块分布_持股市值.xlsx\n')


# In[18]:


####求出流通市值占比
流通市值占比 = pd.DataFrame({'流通市值占比':['创业板','中小板','主板','科创板','合计','创业板','中小板','主板','科创板'],
                       label:[df_all['创业板'],df_all['中小企业板'],df_all['主板'],
                              df_all['科创板'],sum(df_all),df_all['创业板']/sum(df_all),
                              df_all['中小企业板']/sum(df_all),df_all['主板']/sum(df_all),
                              df_all['科创板']/sum(df_all)]
                      })
流通市值占比.to_excel("./result/板块分布_流通市值占比.xlsx",index=False)
print('已经输出流通市值占比至板块分布_流通市值占比.xlsx\n')


# In[19]:


持股市值占比 = pd.DataFrame({"持股市值占比":['创业板','中小板','主板'],
                       label:[df_持有['创业板']/sum(df_持有),
                              df_持有['中小企业板']/sum(df_持有),
                              df_持有['主板']/sum(df_持有)
                             ]
                      })
持股市值占比.to_excel("./result/板块分布_持股市值占比.xlsx",index=False)
print('已经输出持股市值占比至板块分布_持股市值占比.xlsx\n')


# In[ ]:


print("程序结束，谢谢您的使用！\n")

