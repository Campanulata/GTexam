from os.path import split
from typing import Pattern
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re


document = docx.Document('/Users/tylor/Desktop/1.docx') 
par = document.paragraphs
for i in par:
    i.text=i.text.replace(r'__','')
    i.text=i.text.replace('μ',r'\mu')
    

# 导出图片到img文件夹
for rel in document.part._rels:
    rel = document.part._rels[rel]               #获得资源
    if "image" not in rel.target_ref:
        continue
    imgName = re.findall("/(.*)",rel.target_ref)[0]
    with open('img' + "/" + imgName,"wb") as f:
        f.write(rel.target_part.blob)
# 替换空行，此方法库中没实现
for i in range(len(par)):#清空非必要信息
    if par[i].text.find("百校翻联监")>=0 or par[i].text.find("注意事项")>=0:
        par[i].clear()
for i in range(len(par)):#清空空行
    if par[i].text=='' and i<len(par)-1:#第par[i]是空的
        t=i
        while par[t].text=='' and t<len(par)-1:
            t+=1
        diff=t-i
        for j in range(t,len(par)):
            par[j-diff]=par[j]
# 手动配置信息
listWithImg=[1,3,4,6,7,9,11,12,13,18]
maxNum = 18
choiceNum = 12
#
dictABCD = {'A': '', 'B': '', 'C': '','D':''}
for item in ['A','B','C','D']:
    dictABCD[item]=item+' .'#docx中选项格式
#
listQuestion = [0]
listABCD = [0]



for n in range(1, maxNum + 1):  # 当前为第n题，得到题行号
    for i in range(listQuestion[n - 1], len(par)):
        no = str(n) + ' .'
        if par[i].text.find(no) >= 0:
            listQuestion.append(i)
            break
listQuestion.append(len(par)-1)
for n in range(1, choiceNum + 1):  # 当前为第n题，得到A选项行号
    for i in range(listQuestion[n + 1], listQuestion[n], -1):
        if par[i].text.find(dictABCD['A']) >= 0:
            listABCD.append(i)
            break

print(listABCD)

print(listQuestion)
for i in listQuestion:
    print(par[i].text)
# 选择题
choiceAll = ''
for i in range(1, choiceNum + 1):  # 生成选择题
    # 题干
    choice1 = r'\question[6] '
    for j in range(listQuestion[i], listABCD[i]):
        choice1 += par[j].text.lstrip()
    if i in listWithImg:
        choice1+='\n'+r'\begin{center}'+'\n'+r'\includegraphics[]{img/image'+str(listWithImg[0])+r'.jpeg}'+'\n'+r'\end{center}'+'\n'
        listWithImg.remove(i)
    # 选项
    choice2 = ''  
    choiceA = listABCD[i]
    choiceAD = listQuestion[i + 1] - listABCD[i]
    if choiceAD == 4:
        choice2 = r'\fourchoices{' + par[choiceA].text.lstrip() + r'}{' + par[choiceA + 1].text.lstrip() + r'}{' + par[choiceA + 2].text.lstrip() + r'}{' + par[choiceA + 3].text.lstrip() + r'}'
    if choiceAD==2:
        ab=par[choiceA].text
        cd=par[choiceA+1].text
        choice2 = r'\fourchoices{' + ab.split(dictABCD['B'],1)[0].lstrip() + r'}{' + ab.split(dictABCD['B'],1)[1].lstrip() + r'}{' + cd.split(dictABCD['B'],1)[0].lstrip() + r'}{' + cd.split(dictABCD['C'],1)[1].lstrip()+ r'}'

    #尾部
    choice3 = r'\begin{solution}{4cm}' + '\n'+'\n' + r'\end{solution}' + '\n' + '\n'+ '\n'+ '\n'
    choicet = choice1 +'\n'+ choice2 +'\n'+ choice3
    choiceAll += choicet
# 标准化ABCD选项
for item in dictABCD:
    choiceAll=choiceAll.replace(dictABCD[item],'')

# 清楚空格
choiceAll=choiceAll.replace(' ','')
# tex标准化
choiceAll=choiceAll.replace(r'\question[6]','\question[6] ')
# 添加$
listReplace=['v_0','F_0','F_N','t_0','kg·m^2','kg/m^2','t_b','t_a','10m/s^2','g=10m/s^2',r'\mu',r'\nu',r't_{1}',r'^{2}',r's_{0}',r'\cdots','x_1','x_3','x_4','x_5','x_s','v_4','_m/s^2','_m/s','m_1','m_2','_1']
for i in listReplace:
    choiceAll=choiceAll.replace(i,'$'+i+'$')
# frac添加$
choiceAll=choiceAll.replace(r'\frac',r'$\frac')
tmp=choiceAll.count(r'\frac',0)
t=0
listStr = list(choiceAll)
for i in range(tmp):
    print(choiceAll.find(r'\frac',t))
    t=choiceAll.find(r'\frac',t)+1
    times=0
    n=t
    while (listStr[n] !=r'{'):
        n+=1
    times+=1
    while (listStr[n] !=r'}'):
        n+=1
    times+=1
    while (listStr[n] !=r'{'):
        n+=1
    times+=1
    while (listStr[n] !=r'}'):
        n+=1
    times+=1
    listStr.insert(n+1,r'$')
choiceAll=''
for i in listStr:
    choiceAll+=i

# 非选择题
unChoiceAll=''
for i in range(choiceNum + 1, maxNum + 1):  # 生成选择题
    # 题干
    choice1 = ''
    for j in range(listQuestion[i], listQuestion[i+1]):
        choice1 += par[j].text.lstrip()
    choice1=choice1.replace(' ','')
    result=re.findall(r'[^\u4E00-\u9FA5]+',choice1)
    result=list(set(result))
    for m in range(0,len(result)-1):
        for p in range(0,len(result)-1-m):
            if len(result[p])<len(result[p+1]):
                result[p],result[p+1]=result[p+1],result[p]
    for j in ['(1)','(2)']:
        if j in result:
            result.remove(j)
    
    for j in range(0,len(result)):
        if len(result[j])>1:
            resultT=result[j].replace('x','x ')
            resultT=result[j].replace('v','v ')
            resultT=result[j].replace('a','a ')
            
            choice1=choice1.replace(result[j],'$'+resultT+'$')
    choice1 = r'\question[6]'+choice1
    if i in listWithImg:
        choice1+='\n'+r'\begin{center}'+'\n'+r'\includegraphics[]{img/image'+str(listWithImg[0])+r'.jpeg}'+'\n'+r'\end{center}'+'\n'
        listWithImg.remove(i)
    unChoiceAll+=choice1+'\n'
# 去除空格
unChoiceAll=unChoiceAll.replace(' ','')
unChoiceAll=unChoiceAll.replace(r'$$','')
# tex标准化
unChoiceAll=unChoiceAll.replace(r'\question[6]','\question[6] ')
dictSub={'1':'','2':'','3':'','4':''}
for i in dictSub:
    dictSub[i]='('+i+')'
for i in dictSub:
    unChoiceAll=unChoiceAll.replace(dictSub[i],'\n'+dictSub[i])


# # $
# for i in listReplace:
#     unChoiceAll=unChoiceAll.replace(i,'$'+i+'$')
# # frac
# unChoiceAll=unChoiceAll.replace(r'\frac',r'$\frac')
# tmp=unChoiceAll.count(r'\frac',0)
# t=0
# listStr = list(unChoiceAll)
# for i in range(tmp):
#     print(unChoiceAll.find(r'\frac',t))
#     t=unChoiceAll.find(r'\frac',t)+1
#     times=0
#     n=t
#     while (listStr[n] !=r'{'):
#         n+=1
#     times+=1
#     while (listStr[n] !=r'}'):
#         n+=1
#     times+=1
#     while (listStr[n] !=r'{'):
#         n+=1
#     times+=1
#     while (listStr[n] !=r'}'):
#         n+=1
#     times+=1
#     listStr.insert(n+1,r'$')
# unChoiceAll=''
# for i in listStr:
#     unChoiceAll+=i
with open("test.txt", "w") as f:
    f.write(unChoiceAll)
