import jieba
import nltk
import pandas as pd
import numpy as np
import xlwt
import wordcloud
from collections import Counter
import matplotlib.pyplot as matplo

pd.set_option('display.max_columns', 5)
pd.set_option('display.max_rows', 250)
pd.set_option('display.width', 50)


class DataProcess:
    def __init__(self, xls, sheet):
        self.xls = xls
        self.data = pd.read_excel(xls, sheet_name=sheet)

    def drop_blank(self, coloum_name):  # 删除存在的空行
        self.data = self.data.dropna(subset=coloum_name)
        self.data.reset_index(drop=True, inplace=True)
        print("已完成过滤！")
        return

    # def average_score(self):#计算评分平均数和有用的平均数
    #     gradelist = ['很差', '较差', '还行', '推荐', '力荐']
    #     sample = list(self.data['score'].values)
    #     sum = 0.0
    #     sum2 = 0.0
    #     for i in sample:
    #         for j in gradelist:
    #             if i == j:
    #                 sum += (gradelist.index(j)+1)
    #             else:
    #                 continue
    #     avg1 = sum/len(sample)
    #     for i in list(self.data['usefulNum'].values):
    #         sum2 += i
    #     avg2 = sum2/len(sample)
    #     print("评分的平均值为:",avg1)
    #     print("有用的平均值为:", avg2)
    #     return

    def find_useful(self):  # 最有用的
        self.data.sort_values(by='usefulnum')
        s = self.data.loc[0]
        print("有用数最高的数据是:\n{:}".format(s))
        return

    def word_cloud(self):  # 绘制词云
        list2 = list(self.data['comment'])
        ls = []
        for i in list2:
            i2 = jieba.cut(str(i))
            for j in i2:
                if len(j) > 1:
                    ls.append(j)
        top3 = Counter(ls).most_common(3)  # 高频词统计
        print(top3)
        print("出现最多的三个词是：")
        for i in top3:
            print(i[0] + ":出现了{}次".format(i[1]))
            txt = ' '.join(ls)
            input("按任意键绘制词云")
            w = wordcloud.WordCloud(font_path="msyh.ttc",
                                    width=500, height=350,
                                    background_color='white',
                                    max_words=40
                                    )
            w.generate(txt)
            matplo.figure("词云")
            matplo.imshow(w)
            matplo.axis("off")
            matplo.show()
            return

    def word_analyse(self):
        list2 = list(self.data['comment'])
        ls = []
        CC = []
        CD = []
        DT = []
        EX = []
        FW = []
        IN = []
        JJ = []
        JJR = []
        JJS = []
        LS = []
        MD = []
        NN = []
        NNS = []
        NNP = []
        NNPS = []
        PDT = []
        POS = []
        PRP = []
        PRP2 = []
        RB = []
        RBR = []
        RBS = []
        RP = []
        SYM = []
        TO = []
        UH = []
        VB = []
        VBD = []
        VBG = []
        VBN = []
        VBP = []
        VBZ = []
        WDT = []
        WP = []
        WPS = []
        WRB = []
        OTH =[]
        for line in list2:
            tokens = nltk.word_tokenize(line)
            pos_tags = nltk.pos_tag(tokens)
            for word, pos in pos_tags:
                if pos == 'CC':
                    CC.append(word)
                elif pos == 'CD':
                    CD.append(word)
                elif pos == 'DT':
                    DT.append(word)
                elif pos == 'EX':
                    NNS.append(word)
                elif pos == 'FW':
                    FW.append(word)
                elif pos == 'IN':
                    IN.append(word)
                elif pos == 'JJ':
                    JJ.append(word)
                elif pos == 'JJR':
                    JJR.append(word)
                elif pos == 'JJS':
                    JJS.append(word)
                elif pos == 'LS':
                    LS.append(word)
                elif pos == 'MD':
                    MD.append(word)
                elif pos == 'NN':
                    NN.append(word)
                elif pos == 'NNS':
                    NNS.append(word)
                elif pos == 'NNP':
                    NNP.append(word)
                elif pos == 'NNPS':
                    NNPS.append(word)
                elif pos == 'PDT':
                    PDT.append(word)
                elif pos == 'POS':
                    POS.append(word)
                elif pos == 'PRP':
                    PRP.append(word)
                elif pos == 'PRP$':
                    PRP2.append(word)
                elif pos == 'RB':
                    RB.append(word)
                elif pos == 'RBR':
                    RBR.append(word)
                elif pos == 'RBS':
                    RBS.append(word)
                elif pos == 'RP':
                    RP.append(word)
                elif pos == 'SYM':
                    SYM.append(word)
                elif pos == 'TO':
                    TO.append(word)
                elif pos == 'UH':
                    UH.append(word)
                elif pos == 'VB':
                    VB.append(word)
                elif pos == 'VBD':
                    VBD.append(word)
                elif pos == 'VBG':
                    VBG.append(word)
                elif pos == 'VBN':
                    VBN.append(word)
                elif pos == 'VBP':
                    VBP.append(word)
                elif pos == 'VBZ':
                    VBZ.append(word)
                elif pos == 'WDT':
                    WDT.append(word)
                elif pos == 'WP':
                    WP.append(word)
                elif pos == 'WP$':
                    WPS.append(word)
                elif pos == 'WRB':
                    WRB.append(word)
                else:
                    OTH.append(word)
        print("词性以及符合词性的个数如下:")
        print('Coordinating conjunction:', len(CC))
        print('Cardinal number:', len(CD))
        print('Determiner:', len(DT))
        print('Existential there:', len(EX))
        print('Foreign word:', len(FW))
        print('Preposition or subordinating conjunction:', len(IN))
        print('Adjective:', len(JJ))
        print('Adjective, comparative:', len(JJR))
        print('Adjective, superlative:', len(JJS))
        print('List item marker:', len(LS))
        print('Modal:', len(MD))
        print('Noun, singular or mass:', len(NN))
        print('Noun, plural:', len(NNS))
        print('Proper noun, singular:', len(NNP))
        print('Proper noun, plural:', len(NNPS))
        print('Predeterminer:', len(PDT))
        print('Possessive ending:', len(POS))
        print('Personal pronoun:', len(PRP))
        print('Possessive pronoun:', len(PRP2))
        print('Adverb:', len(RB))
        print('Adverb, superlative:', len(RBS))
        print('Particle:', len(RP))
        print('Symbol:', len(SYM))
        print('to:', len(TO))
        print('Interjection:', len(UH))
        print('Verb, base form:', len(VB))
        print('Verb, past tense:', len(VBD))
        print('Verb, gerund or present participle:', len(VBG))
        print('Verb, past participle:', len(VBN))
        print('Verb, non-3rd person singular present:', len(VBP))
        print('Wh-determiner:', len(WDT))
        print('Wh-pronoun:', len(WP))
        print('Possessive wh-pronoun:', len(WPS))
        print('Wh-adverb:', len(WRB))






                     # print(word, pos)


    def save_to_excel(self):
        self.data.to_excel('final.xls', sheet_name='0')
        print('数据已保存')
        return


obj = DataProcess(r'D:\Library\Code\PyCharm_Project\douban\venv\text.xls', sheet=0)  # 创建处理对象
obj.drop_blank('ranking')  # 清除'ranking'中的空值行
obj.drop_blank('usefulnum')  # 清除'usefulNum'中的空值行
obj.drop_blank('date')  # 清除'date'中的空值行
obj.drop_blank('comment')  # 清除'comment'中的空值行
obj.find_useful()  # 找出最有用的 评论
obj.word_cloud()  # 词云
obj.word_analyse()
obj.save_to_excel()  # 保存至excel
