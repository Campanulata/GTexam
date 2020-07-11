from os.path import split
from typing import Pattern
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re
def replaceStr(allStr,str1,str2):
    allStr=allStr.replace(str1,str2)
def delQuad(allStr):
    return allStr.replace(' ','')
def latex(allStr):
    result=re.findall(r'[^\u4E00-\u9FA5]+',allStr)
    result=list(set(result))
    for j in range(len(result)):
        if len(result[j])<3 :
            result[j]='willDel'
        if r'includegraphics' in result[j] or r'image' in result[j]:
            result[j]='willDel'
    while 'willDel' in result:
        result.remove('willDel')
    for m in range(0,len(result)-1):#从大到小排序
        for p in range(0,len(result)-1-m):
            if len(result[p])<len(result[p+1]):
                result[p],result[p+1]=result[p+1],result[p]
    #移除不需要加$的字符
    for j in ['(1)','(2)','(3)','(4)','(a)','(b)',r'\begin{center}\includegraphics[',r'{img/image1.png}\end{center}',r'\begin{center}\includegraphics[]{img/image1.png}\end{center}',r'.\begin{center}\includegraphics[',r'{img/image1.png}\end{center}A.',r'?\begin{center}\includegraphics[']:
        if j in result:
            result.remove(j)
    for j in range(0,len(result)):#替换 加$
        resultT=' '.join(re.compile('.{1,1}').findall(result[j]))
        allStr=allStr.replace(result[j],'$'+resultT+'$')
    # tex标准化
    allStr=allStr.replace(r'\question[6]','\question[6] ')
    return allStr

document = docx.Document('/Users/tylor/Desktop/1.docx') 
par = document.paragraphs
# 导出图片到img文件夹
imgCount=0
for rel in document.part._rels:
    rel = document.part._rels[rel]               #获得资源
    if "image" not in rel.target_ref:
        continue
    imgName = re.findall("/(.*)",rel.target_ref)[0]
    imgCount+=1
    with open('img' + "/" + imgName,"wb") as f:
        f.write(rel.target_part.blob)
#加图
for j in range(1,imgCount+1):#imagej.png
    for i in par:
        if i.text=='':
            i.text=r'\begin{center}'+r'\includegraphics[]{img/image'+str(j)+r'.png}'+r'\end{center}'
            break
        continue
for i in par:
    i.text=i.text.replace(r'∠',r'\angle')
    i.text=i.text.replace(r'ⅱ',r'ii')
    
    for j in ['\key{}','。','、',',',':','(i)','(ii)','(',')',']','，']:
        i.text=i.text.replace(j,'答案'+j+'答案')
    


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
listWithImg=[]
maxNum = 14
choiceNum = 8
#
dictABCD = {'A': '', 'B': '', 'C': '','D':''}
for item in ['A','B','C','D']:#docx中选项格式
    dictABCD[item]=item+' .'
#
listQuestion = [0]
listABCD = [0]
# 当前为第n题，得到题行号
for n in range(1, maxNum + 1):  
    for i in range(listQuestion[n - 1], len(par)):
        no = str(n) + ' .'
        if par[i].text.find(no) >= 0:
            listQuestion.append(i)
            break
listQuestion.append(len(par)-1)
# 当前为第n题，得到A选项行号
for n in range(1, choiceNum + 1):  
    for i in range(listQuestion[n + 1], listQuestion[n], -1):
        if par[i].text.find(dictABCD['A']) >= 0:
            listABCD.append(i)
            break
print(listABCD)
print(listQuestion)
# 选择题
choiceAll = ''
imgNo=1
for i in range(1, choiceNum + 1):  # 生成选择题
    # 题干
    choice1 = ''
    for j in range(listQuestion[i], listABCD[i]):
        par[j].text.replace(' ','')
        choice1 += par[j].text.lstrip()
    choice1=choice1[3:]
    choice1=delQuad(choice1)
    choice1=latex(choice1)
    choice1=r'\question[6] '+choice1

    # 选项
    choice2 = ''  
    choiceA = listABCD[i]
    a=par[choiceA].text.lstrip().replace(dictABCD['A'],'')
    b=par[choiceA+1].text.lstrip().replace(dictABCD['B'],'')
    c=par[choiceA+2].text.lstrip().replace(dictABCD['C'],'')
    d=par[choiceA+3].text.lstrip().replace(dictABCD['D'],'')
    a=delQuad(a)
    b=delQuad(b)
    c=delQuad(c)
    d=delQuad(d)
    choice2 = r'\fourchoices{' + latex(a) + r'}{' + latex(b) + r'}{' + latex(c) + r'}{' + latex(d) + r'}'


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
    choice1=choice1.replace(' ','').split('.',1)[1]

    #
    choice1=delQuad(choice1)
    choice1=latex(choice1)
    choice1 = r'\question[6]'+choice1

    unChoiceAll+=choice1+'\n'
# 去除空格

unChoiceAll=unChoiceAll.replace(r'$$','')


def latexQuad(strAll):
    for i in [r'\times',r'\pi',r'\Delta',r'\rightarrow',r'\angle',r'\leqslant']:
        strAll=strAll.replace(i,i+' ')
    return strAll

choiceAll=choiceAll.replace(' ','')
for i in dictABCD:
    choiceAll=choiceAll.replace(dictABCD[i],'')
unChoiceAll=unChoiceAll.replace(' ','')
choiceAll=latexQuad(choiceAll)
unChoiceAll=latexQuad(unChoiceAll)
unChoiceAll=unChoiceAll.replace(r'\question[6]',r'\question[6] ')
#中文保护
def delHeadFoot(strAll,key):
    return strAll.replace('答案'+key+'答案',key)
#中文保护 替换
for i in ['。','、',',',':','(i)','(ii)','(',')',']','，']:
    choiceAll=delHeadFoot(choiceAll,i)
    unChoiceAll=delHeadFoot(unChoiceAll,i)
choiceAll=choiceAll.replace(r'答案$\key{}$答案',r'\key{}')
unChoiceAll=unChoiceAll.replace(r'答案$\key{}$答案',r'\key{}')

#sub
dictSub={'1':'','2':'','3':'','4':'','5':'','6':'','7':''}
for i in dictSub:#生产sub标题
    dictSub[i]='('+i+')'
for i in dictSub:#sub加\n
    unChoiceAll=unChoiceAll.replace(dictSub[i],'\n\n'+dictSub[i])
with open("Part1Choice.tex", "w") as f:
    f.write(choiceAll)
with open("Part2UnChoice.tex", "w") as f:
    f.write(unChoiceAll)
