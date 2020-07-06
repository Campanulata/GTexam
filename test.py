from os.path import split
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re


document = docx.Document('/Users/tylor/Desktop/1.docx') 
par = document.paragraphs
# 导出图片到img文件夹
for rel in document.part._rels:
    rel = document.part._rels[rel]               #获得资源
    if "image" not in rel.target_ref:
        continue
    imgName = re.findall("/(.*)",rel.target_ref)[0]
    with open('img' + "/" + imgName,"wb") as f:
        f.write(rel.target_part.blob)
# 替换空行，此方法库中没实现
for i in range(len(par)):
    if par[i].text.find("百校翻联监")>=0:
        par[i].clear()
for k in range(111):
    for i in range(len(par)):
        if par[i].text=='':
            for j in range(i,len(par)-1):
                par[j]=par[j+1]
# 手动配置信息
listWithImg=[1,3,4,6,7,9,11,12]
maxNum = 18
choiceNum = 12
listQuestion = [0]
listABCD = [0]
dictABCD = {'A': '', 'B': '', 'C': '','D':''}
for item in ['A','B','C','D']:
    dictABCD[item]=item+' .'


for n in range(1, maxNum + 1):  # 当前为第n题，得到题行号
    for i in range(listQuestion[n - 1], len(par)):
        no = str(n) + ' .'
        if par[i].text.find(no) >= 0:
            listQuestion.append(i)
            break

for n in range(1, choiceNum + 1):  # 当前为第n题，得到选项行号
    for i in range(listQuestion[n + 1], listQuestion[n], -1):
        if par[i].text.find(dictABCD['A']) >= 0:
            listABCD.append(i)
            break

print(listABCD)

print(listQuestion)
for i in listQuestion:
    print(par[i].text)

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

for item in dictABCD:
    choiceAll=choiceAll.replace(dictABCD[item],'')
choiceAll=choiceAll.replace(' ','')
choiceAll=choiceAll.replace(r'\question[6]','\question[6] ')
listReplace=['v_0','F_0','F_N','t_0','kg·m^2','kg/m^2','t_b','t_a','10m/s^2','g=10m/s^2',r'\mu']
for i in listReplace:
    choiceAll=choiceAll.replace(i,'$'+i+'$')
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
with open("test.txt", "w") as f:
    f.write(choiceAll)
