from os import replace
from os.path import split
from typing import Pattern
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re

class Paper:
    par=''
    document=''
    choiceAll = ''
    unChoiceAll=''
    imgCount=0
    maxNum = 0
    choiceNum = 0
    listQuestion = [0]
    listABCD = [0]
    dictUnLatexProtect=['\key{}','。','、',',',':','(',')',']','，','（','）']
    dictABCD = {'A': '', 'B': '', 'C': '','D':''}
    dictSub={'1':'','2':'','3':'','4':'','5':'','6':'','7':''}
    dictLatexQuad=[r'\times',r'\pi',r'\Delta',r'\rightarrow',r'\angle',r'\leqslant',r'\cdot',r'\question[6]']
    dictIrrelevant=['百校翻联监','注意事项']
    dictReplace={   r'∠':r'\angle',
                    r'ⅱ':r'\romannumeral2',
                    r' ':'',
                    r'Ⅰ':r'\uppercase\expandafter{\romannumeral1}'}
    def delQuad(self,allStr):
        return allStr.replace(' ','')
    def latex(self,allStr):
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
    def get_ABCD_adn_sub(self):
        for item in ['A','B','C','D']:#docx中选项格式
            self.dictABCD[item]=item+' .'
        for i in self.dictSub:#生产sub标题
            self.dictSub[i]='（'+i+'）'
    def get_list_question(self):
        for n in range(1, self.maxNum + 1):  
            for i in range(self.listQuestion[n - 1], len(self.par)):
                no = str(n) + ' .'
                if self.par[i].text.find(no) >= 0:
                    self.listQuestion.append(i)
                    break
        self.listQuestion.append(len(self.par))
    def get_list_ABCD(self):
        for n in range(1, self.choiceNum + 1):  
            for i in range(self.listQuestion[n + 1], self.listQuestion[n], -1):
                if self.par[i].text.find(self.dictABCD['A']) >= 0:
                    self.listABCD.append(i)
                    break
    def image_to_img(self):
        for rel in self.document.part._rels:
            rel = self.document.part._rels[rel]      #获得资源
            if "image" not in rel.target_ref:
                continue
            imgName = re.findall("/(.*)",rel.target_ref)[0]
            self.imgCount+=1
            with open('img' + "/" + imgName,"wb") as f:
                f.write(rel.target_part.blob)
    def add_image(self):
        for j in range(1,self.imgCount+1):#imagej.png
            for i in self.par:
                if i.text=='':
                    i.text=r'\begin{center}'+r'\includegraphics[]{img/image'+str(j)+r'.png}'+r'\end{center}'
                    break
                continue
    def for_i_in_par(self):
        for i in self.par:
            for j in self.dictReplace:
                i.text=i.text.replace(j,self.dictReplace[j])            
            for j in self.dictUnLatexProtect:
                i.text=i.text.replace(j,'答案'+j+'答案')
    def irrelevant_information(self):
        for i in range(len(self.par)):#清空非必要信息
            for j in self.dictIrrelevant:
                if self.par[i].text.find(j)>=0 :
                    self.par[i].clear()
    def del_empty_line(self):
        for i in range(len(self.par)):#清空空行
            if self.par[i].text=='' and i<len(self.par)-1:#第par[i]是空的
                t=i
                while self.par[t].text=='' and t<len(self.par)-1:
                    t+=1
                diff=t-i
                for j in range(t,len(self.par)):
                    self.par[j-diff]=self.par[j]
    def get_choice_all(self):
        for i in range(1, self.choiceNum + 1):  # 生成选择题
            # 题干
            choice1 = ''
            for j in range(self.listQuestion[i], self.listABCD[i]):
                self.par[j].text.replace(' ','')
                choice1 += self.par[j].text.lstrip()
            choice1=choice1[2:]
            choice1=self.delQuad(choice1)
            choice1=choice1+r'答案\key{}答案'
            choice1=self.latex(choice1)
            choice1=r'\question[6] '+choice1

            # 选项
            choice2 = ''  
            choiceA = self.listABCD[i]
            a=self.par[choiceA].text.lstrip().replace(self.dictABCD['A'],'')[2:]
            b=self.par[choiceA+1].text.lstrip().replace(self.dictABCD['B'],'')[2:]
            c=self.par[choiceA+2].text.lstrip().replace(self.dictABCD['C'],'')[2:]
            d=self.par[choiceA+3].text.lstrip().replace(self.dictABCD['D'],'')[2:]
            a=self.delQuad(a)
            b=self.delQuad(b)
            c=self.delQuad(c)
            d=self.delQuad(d)
            choice2 = r'\fourchoices{' + self.latex(a) + r'}{' + self.latex(b) + r'}{' + self.latex(c) + r'}{' + self.latex(d) + r'}'


            #尾部
            choice3 = r'\begin{solution}{4cm}' + '\n'+'\n' + r'\end{solution}' + '\n' + '\n'+ '\n'+ '\n'
            choicet = choice1 +'\n'+ choice2 +'\n'+ choice3
            self.choiceAll += choicet
    def get_unchoice_all(self):
        for i in range(self.choiceNum + 1, self.maxNum + 1):  # 生成选择题
            # 题干
            choice1 = ''
            for j in range(self.listQuestion[i], self.listQuestion[i+1]):
                choice1 += self.par[j].text.lstrip()
            # sub
            choice1=choice1.replace(' ','').split('.',1)[1]

            #
            choice1=self.delQuad(choice1)
            choice1=self.latex(choice1)
            choice1 = r'\question[6]'+choice1

            self.unChoiceAll+=choice1+'\n'
        # 去除空格

        self.unChoiceAll=self.unChoiceAll.replace(r'$$','')
    def latex_will_work(self):

        # 去ABCD
        # for i in self.dictABCD:
        #     self.choiceAll=self.choiceAll.replace(self.dictABCD[i],'')

        # 去空格
        self.choiceAll=self.choiceAll.replace(' ','')
        self.unChoiceAll=self.unChoiceAll.replace(' ','')
        #
        self.choiceAll=self.choiceAll.replace(r'\begin{center}',r'\key{}\begin{center}')
        self.choiceAll=self.choiceAll.replace(r'\end{center}答案$\key{}$答案',r'\end{center}')
        self.choiceAll=self.choiceAll.replace(r'\end{center}\key{}',r'\end{center}')
        
        # 加空格
        for i in self.dictLatexQuad:
            self.choiceAll=self.choiceAll.replace(i,i+' ')
            self.unChoiceAll=self.unChoiceAll.replace(i,i+' ')
        # 删除中文保护
        self.choiceAll=self.choiceAll.replace(r'答案$\key{}$答案',r'\key{}')
        for i in self.dictUnLatexProtect:
            self.choiceAll=self.choiceAll.replace('答案'+i+'答案',i)
            self.unChoiceAll=self.unChoiceAll.replace('答案'+i+'答案',i)
        # sub加回车
        for i in self.dictSub:#sub加\n
            self.unChoiceAll=self.unChoiceAll.replace(self.dictSub[i],'\n\n'+self.dictSub[i])
    def write_to_tex(self):
        with open("Part1Choice.tex", "w") as f:
            f.write(self.choiceAll)
        with open("Part2UnChoice.tex", "w") as f:
            f.write(self.unChoiceAll)
