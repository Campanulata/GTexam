from os.path import split
from typing import Pattern
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re

def latex(allStr):
    allStr=allStr.replace(" ",'')
    result=re.findall(r'[^\u4E00-\u9FA5]+',allStr)
    result=list(set(result))
    for m in range(0,len(result)-1):#从大到小排序
        for p in range(0,len(result)-1-m):
            if len(result[p])<len(result[p+1]):
                result[p],result[p+1]=result[p+1],result[p]
    for j in range(len(result)):
        if len(result[j])<3:
            result[j]='dsadasdas'
        if '\n' in result[j]:
            result[j]='dsadasdas'
    for j in ['(1)','(2)']:#移除不需要加$的字符
        if j in result:
            result.remove(j)
    for j in range(0,len(result)):#替换 加$
        resultT=result[j].replace('x','x ')
        resultT=result[j].replace('v','v ')
        resultT=result[j].replace('a','a ')
        allStr=allStr.replace(result[j],'$'+resultT+'$')
    # tex标准化
    allStr=allStr.replace(r'\question[6]','\question[6] ')
    return allStr

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

# 选择题
choiceAll = ''
for i in range(1, choiceNum + 1):  # 生成选择题
    # 题干
    choice1 = ''
    for j in range(listQuestion[i], listABCD[i]):
        par[j].text.replace(' ','')
        choice1 += par[j].text.lstrip()
    choice1=latex(choice1)
    choice1=r'\question[6] '+choice1
    if i in listWithImg:
        choice1+='\n'+r'\begin{center}'+'\n'+r'\includegraphics[]{img/image'+str(listWithImg[0])+r'.jpeg}'+'\n'+r'\end{center}'+'\n'
        listWithImg.remove(i)
    # 选项
    choice2 = ''  
    choiceA = listABCD[i]
    choiceAD = listQuestion[i + 1] - listABCD[i]
    if choiceAD == 4:
        choice2 = r'\fourchoices{' + latex(par[choiceA].text.lstrip()) + r'}{' + latex(par[choiceA + 1].text.lstrip()) + r'}{' + latex(par[choiceA + 2].text.lstrip()) + r'}{' + latex(par[choiceA + 3].text.lstrip()) + r'}'
    if choiceAD==2:
        ab=par[choiceA].text
        cd=par[choiceA+1].text
        a=ab.split(dictABCD['B'],1)[0].lstrip()
        b=ab.split(dictABCD['B'],1)[1].lstrip()
        c=cd.split(dictABCD['C'],1)[0].lstrip()
        d=cd.split(dictABCD['C'],1)[1].lstrip()
        # 标准化ABCD选项
        a=a.replace(dictABCD['A'],'')
        b=b.replace(dictABCD['B'],'')
        c=c.replace(dictABCD['C'],'')
        d=d.replace(dictABCD['D'],'')
        choice2 = r'\fourchoices{' + latex(a) + r'}{' + latex(b) + r'}{' + latex(c) + r'}{' + latex(d)+ r'}'

    #尾部
    choice3 = r'\begin{solution}{4cm}' + '\n'+'\n' + r'\end{solution}' + '\n' + '\n'+ '\n'+ '\n'
    choicet = choice1 +'\n'+ choice2 +'\n'+ choice3
    choiceAll += choicet


# 非选择题
unChoiceAll=''
for i in range(choiceNum + 1, maxNum + 1):  # 生成选择题
    # 题干
    choice1 = ''
    for j in range(listQuestion[i], listQuestion[i+1]):
        choice1 += par[j].text.lstrip()
    # sub
    choice1=choice1.replace(' ','')
    dictSub={'1':'','2':'','3':'','4':''}
    for j in dictSub:
        dictSub[j]='('+j+')'
    for j in dictSub:
        choice1=choice1.replace(dictSub[j],'\n'+dictSub[j])
    #
    choice1=latex(choice1)
    choice1 = r'\question[6]'+choice1
    if i in listWithImg:
        choice1+='\n'+r'\begin{center}'+'\n'+r'\includegraphics[]{img/image'+str(listWithImg[0])+r'.jpeg}'+'\n'+r'\end{center}'+'\n'
        listWithImg.remove(i)
    unChoiceAll+=choice1+'\n'
# 去除空格

unChoiceAll=unChoiceAll.replace(r'$$','')





choiceAll=choiceAll.replace(' ','')
unChoiceAll=unChoiceAll.replace(' ','')
unChoiceAll=unChoiceAll.replace(r'\question[6]',r'\question[6] ')
with open("Part1Choice.tex", "w") as f:
    f.write(choiceAll)
with open("Part2UnChoice.tex", "w") as f:
    f.write(unChoiceAll)
