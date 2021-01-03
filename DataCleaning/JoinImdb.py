import re

import pandas as pd
from fuzzywuzzy import fuzz

import sys  # 导入sys模块
sys.setrecursionlimit(3000)  # 将默认的递归深度修改为3000

# 正则表达，只保留大小写字母和阿拉伯数字
reg = "[^0-9A-Za-z]"
num = "[^0-9]"

imdbData = pd.read_excel(io="D:\\Projects\\PythonProjects\\DataCleaning\\data\\imdbTest.xlsx",
                         index_col=1,
                         engine='openpyxl')

amazonData = pd.read_excel(io="D:\\Projects\\PythonProjects\\DataCleaning\\data\\Step1Result.xlsx",
                           index_col=0,
                           engine='openpyxl')

mergeData = pd.merge(amazonData, imdbData, on='asin', how="inner")
mergeData.sort_values(by=['电影名', '导演'])
mergeData.reset_index(inplace=True)
unknownList = ["unknown", "unkn", "---", "none"]

mergeRows = len(mergeData)
# 合并主电影
for i in range(mergeRows):
    isNum = False
    if re.sub(num, '', str(mergeData.iloc[i, 1])) == mergeData.iloc[i, 1]:
        isNum = True
    myTitle = re.sub(reg, '', str(mergeData.iloc[i, 1])).lower()
    compareStart = i - 40  # 起始比较行
    if compareStart < 0:
        compareStart = 0
    master = True  # 默认是主电影
    for cmp in range(compareStart, i):
        # 判断是否是master
        if mergeData.iloc[cmp, 12] == "master":
            pass
        else:
            continue
        # 判断至少其中之一isMovie
        if (mergeData.iloc[cmp, 14] == 1) or (mergeData.iloc[i, 14] == 1):
            pass
        else:
            continue

        cmpTitle = re.sub(reg, '', str(mergeData.iloc[cmp, 1])).lower()
        # 判断title的模糊识别
        if isNum or re.sub(num, '', str(mergeData.iloc[cmp, 1])) == mergeData.iloc[cmp, 1]:
            if myTitle == cmpTitle:
                myProducers = str(mergeData.iloc[i, 4]).split(', ')
                cmpProducers = str(mergeData.iloc[cmp, 4]).split(', ')
                isOne = False
                # title相等，逐一判断导演
                for myProducer in myProducers:
                    for cmpProducer in cmpProducers:
                        if fuzz.partial_ratio(myProducer.lower(), cmpProducer.lower()) > 95 or \
                                fuzz.partial_ratio(cmpProducer.lower(), myProducer.lower()) > 95 or \
                                myProducer.lower() in unknownList or cmpProducer.lower() in unknownList:
                            # 是同一个电影
                            isOne = True
                            if not master:
                                # 本电影已经不是主电影，则将主电影asin号赋值给前者
                                mergeData.iloc[cmp, 12] = mergeData.iloc[i, 12]
                            else:
                                master = False
                                # 将前者asin号赋值给后者
                                mergeData.iloc[i, 12] = mergeData.iloc[cmp, 0]
                            # 确保isMovie为1
                            mergeData.iloc[cmp, 14] = 1
                            mergeData.iloc[i, 14] = 1
                            break
                    if isOne:
                        break
        elif fuzz.partial_ratio(myTitle, cmpTitle) > 98 or fuzz.partial_ratio(cmpTitle, myTitle) > 98:
            myProducers = str(mergeData.iloc[i, 4]).split(', ')
            cmpProducers = str(mergeData.iloc[cmp, 4]).split(', ')
            isOne = False
            # title相等，逐一判断导演
            for myProducer in myProducers:
                for cmpProducer in cmpProducers:
                    if fuzz.partial_ratio(myProducer.lower(), cmpProducer.lower()) > 95 or \
                            fuzz.partial_ratio(cmpProducer.lower(), myProducer.lower()) > 95 or \
                            myProducer.lower() in unknownList or cmpProducer.lower() in unknownList:
                        # 是同一个电影
                        isOne = True
                        if not master:
                            # 本电影已经不是主电影，则将主电影asin号赋值给前者
                            mergeData.iloc[cmp, 12] = mergeData.iloc[i, 12]
                        else:
                            master = False
                            # 将前者asin号赋值给后者
                            mergeData.iloc[i, 12] = mergeData.iloc[cmp, 0]
                        # 确保isMovie为1
                        mergeData.iloc[cmp, 14] = 1
                        mergeData.iloc[i, 14] = 1
                        break
                if isOne:
                    break

# 去掉非电影
mergeData.drop(index=mergeData[mergeData['isMovie'].isin([0])].index[0])
# mergeData.drop(index=mergeData[mergeData['isMovie'] == 0].index)
mergeData.drop(columns='isMovie')
mergeData.drop(columns='idx')
# 插入存储asin列
mergeData.insert(loc=13, column="movies", value="none", allow_duplicates=False)
# # 重新计算行数
# mergeRows = len(mergeData)
# # 整理第二次主电影
# for i in range(mergeRows):
#     if mergeData.iloc[i, 12] == "master":
#         mergeData.iloc[i, 15] = mergeData.iloc[i, 0]
#     else:
#         for j in range(i-1, -1, -1):
#             if mergeData.iloc[j, 0] == mergeData.iloc[i, 12]:
#                 mergeData.iloc[j, 15] = mergeData.iloc[j, 15] + "," + mergeData.iloc[i, 0]
#                 break
#
# # 整理第一次非主电影
# amazonRows = len(amazonData)
# amazonData.reset_index(inplace=True)
# movieIter = 0
# # for i in range(amazonRows):
# for i in range(10000):
#     print(i)
#     a = amazonData.iloc[i]
#     if amazonData.iloc[i, 12] == "master":
#         continue
#     else:
#         match = False
#         for j in range(movieIter, mergeRows):
#             a = amazonData.iloc[i, 12]
#             if mergeData.iloc[j, 0] == amazonData.iloc[i, 12]:
#                 if mergeData.iloc[j, 12] == "master":
#                     mergeData.iloc[j, 15] = mergeData.iloc[j, 15] + "," + amazonData.iloc[i, 0]
#                     match = True
#                     movieIter = j
#                     break
#                 else:
#                     for k in range(j - 15, j):
#                         if mergeData.iloc[k, 0] == mergeData.iloc[j, 12]:
#                             mergeData.iloc[k, 15] = mergeData.iloc[k, 15] + "," + amazonData.iloc[i, 0]
#                             break
#             if match:
#                 movieIter = j
#                 break
#         if not match:
#             for j in range(mergeRows):
#                 if mergeData.iloc[j, 0] == amazonData.iloc[i, 12]:
#                     if mergeData.iloc[j, 12] == "master":
#                         mergeData.iloc[j, 15] = mergeData.iloc[j, 15] + "," + amazonData.iloc[i, 0]
#                         match = True
#                         movieIter = j
#                         break
#                     else:
#                         for k in range(j - 15, j):
#                             if mergeData.iloc[k, 0] == mergeData.iloc[j, 12]:
#                                 mergeData.iloc[k, 15] = mergeData.iloc[k, 15] + "," + amazonData.iloc[i, 0]
#                                 match = True
#                                 break
#                 if match:
#                     movieIter = j
#                     break
#
#
# # 去掉非主电影
# mergeData.drop(index=mergeData[mergeData['master'] != "master"].index)

mergeData.to_excel('D:\\Projects\\PythonProjects\\DataCleaning\\data\\MergedData.xlsx',
                   sheet_name='Sheet1',
                   engine='openpyxl')
