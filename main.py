#!/usr/bin/env python
# -*- coding: utf-8 -*-

import random
import os
import xlrd
import xlwt
from xlutils.copy import copy

defaultSheetNameList = tuple('sheet' + str(i) for i in range(25))

def base_dir(filename=None):
    return os.path.join(os.path.dirname(__file__), filename)

def abbreviation_separate(in_word):
    word = in_word
    # print('functin in:', in_word)
    leftBacketBoundary, rightBacketBoundary = word.find('('), word.find(')')
    if leftBacketBoundary <= rightBacketBoundary - leftBacketBoundary - 1 and word[0] == '(':  # 缩写在左, 括号包缩写
        abbreviation = word[1:rightBacketBoundary]
        word = word[rightBacketBoundary + 1:]
    elif leftBacketBoundary <= rightBacketBoundary - leftBacketBoundary - 1 and word[0] != '(':  # 缩写在左, 括号包word
        abbreviation = word[:leftBacketBoundary]
        word = word[leftBacketBoundary + 1:rightBacketBoundary]
    elif leftBacketBoundary > rightBacketBoundary - leftBacketBoundary - 1 and word[0] != '(':  # 缩写在右, 括号包缩写
        abbreviation = word[leftBacketBoundary + 1:rightBacketBoundary]
        word = word[:leftBacketBoundary]
    else:  # 缩写在右，括号包括word
        abbreviation = word[rightBacketBoundary + 1:]
        word = word[leftBacketBoundary + 1:rightBacketBoundary]
    return abbreviation, word

def create_excelfile(filename, sheetnamelist=defaultSheetNameList, filepath=os.path.dirname(os.path.abspath(__file__))):
    newExcelFile = xlwt.Workbook(encoding='utf-8')
    for sheetName in sheetnamelist:
        newExcelFile.add_sheet(sheetName)
    newExcelFile.save(filepath + '\\' + filename + '.xls')

# def sheet_headline(workbook, sheetname, headnamelist):
#     rd = xlrd.open_workbook(workbook)
#     sheet = rd.sheet_by_name(sheetname)
#     wt = copy(rd)  # 复制, xlutils.copy.copy
#
#     sheet1 = wt[sheetname]  # 读取第一个工作表
#     for i in headnamelist:

# def add_tuple_by_sheet_index(workbook, sheetIndex, wordInfo):

#不删除enemies元素，因为这样会导致删除前面的后，后面所有元素下标都减一，所以在一开始可以维护另一个存放所有要背单词的下标，然后在它里面进行random.sample
'''
start_from_record   是否从已有进度上继续，0否1是
killing_num         本次想要背诵数量, 默认是一次性背完这个单词本的所有单词，通过给killing_num赋一个很多大的值来实现
workbookopt         选择哪个单词本
sheetopt            选择哪一个sheet，注意，这里用下标进行选择，你想用名字选取的话，在sheet获取部分选择第二行即可
colopt              通过下标选择背诵情况记录列

excel 结构: [单词， 释义， 剩余背诵次数]

你可以再选择一列用来给用户自己添加笔记，这个功能我没实现，不过很简单

'''
def words_killing(start_from_record=0, killing_num=50000, workbookopt='x.xls', sheetopt=0, colopt=2):
    # 后提示部分
    promptList = ("", ",请注意大小写哦！")  # 用于在用户输入后给予提示(后提示)
    promptOpt = 0   #后提示选择自

    # 前提示部分
    formerPromptList = ("", ",记得缩写和全称都要哦", )    # 用于在用户输入单词前给予适当提示(前提示)
    formerPrompOpt = 0  # 前提示选择子

    # 缩写初始化
    abbreviation = ""

    # 打开选中的单词本
    rd = xlrd.open_workbook(workbookopt)

    # 打开选中的sheet
    sheet = rd.sheet_by_index(sheetopt) #若上面选择用sheet下标进行选择，则用这行
    # sheet = rd.sheet_by_name(sheetopt)  #若上面选择用sheet名进行选择，则用这行

    # 用于最后的保存进度使用
    wt = copy(rd)  # 复制, xlutils.copy.copy
    sheet1 = wt.get_sheet(sheetopt)  # 读取第一个工作表

    # enemies存储需要背的单词[在excel中的下标, 剩余背诵次数]的union
    enemies = []    # 存[excel下标, times]的union
    finshedwords = []     # 存完成的单词

    total = sheet.nrows
    interval = killing_num    # 本次想要背诵数量，目前来看你直接把所有interval替换为killing_num也没问题，只是因为我测试时候用的interval
    start = 1                 # 0行是表属性，因此从excel的1下标行开始读

    # 从已有记录继续
    if start_from_record:
        for line_num in range(start, min(total, start + interval)):
            val = sheet.cell_value(line_num, colopt)
            if val > 0:  #仅将剩余次数大于0的词加入待背诵单词本
                enemies.append([line_num, val])

    # 从新开始背
    else:
        for line_num in range(start, min(total, start + interval)):
            enemies.append([line_num, 1])

            sheet1.write(line_num, colopt, 1)

    # 本次待背诵单词本中单词数量
    row = len(enemies)

    # 在剩余单词数为特定数值时会给予用户提示，以激励其继续背诵
    seprator1 = [i * 10 for i in range(1, 11)]
    seprator2 = [100 + i * 20 for i in range(1, 6)]
    seprator3 = [200 + i * 50 for i in range(1, 5)]
    seprator = [5] + seprator1 + seprator2 + seprator3

    # 用户输入单词
    input_word = ''

    # 退出判断子
    quit = False

    # 通过将单词本打乱实现乱序背诵，乱序背单词的时间复杂度为该库函数时间复杂度，对于后续的回答错误，每次打错换位O(1)
    random.shuffle(enemies)

    # 指向当前背诵单词在enemies中的下标，这里因为单词剩余背诵次数达到0后需要从enemies中去除，所以从后往前进行
    # 至于用户没有一次就答对的单词，会将其与前面的某个单词调换位置，然后不删除，具体见后面实现
    pointer = row - 1

    # 当前背到第ptr个单词，如果想，也可以用killing_num - pointer实现，不过这样可能会因为用户输入killing_num大于选中单词本的选中sheet
    # 中的单词总数而出现问题
    ptr = 1

    while row != 0 and not quit:
        # 注意excel第一行为属性，但是程序中下标从0开始，所以一切与sheet有关的下标均使用excelid，简称ei;与enemies有关的下标为ri
        ei = enemies[pointer][0]

        # word为当前背诵的单词
        word = sheet.cell_value(ei, 0)

        if '(' in word: #有缩写
            formerPrompOpt = 1
            abbreviation, word = abbreviation_separate(word)
        else:   #无缩写
            formerPrompOpt = 0

        # 打印 单词 与 前提示
        print(ptr, '\t' + sheet.cell_value(ei, 1) + formerPromptList[formerPrompOpt])

        # 用于记录用户回答情况，若time == 0，代表用户一遍就答对，否则每次打错time + 1，直到time == 3，直接给出答案并且暂时跳过该词
        time = 0
        if formerPrompOpt == 1: #有缩写
            while time < 3:
                input_word = input('请输入拼写: ')
                if input_word == '1':   # 代表会这个词，直接跳过
                    print('(' + abbreviation + ')' + word) #打印答案，不想要可以去掉。你如果想，可以添加将上一个单词重新加入的功能
                    break
                elif input_word == 'q': #退出背单词，记录依然会被保持
                    order = input('您确定要退出吗[y/n]?')
                    if order == 'y':
                        quit = True
                        break

                # 用户输入的 缩写 和 单词，你也可以在abbreviation_separate函数中添加用户没有给出缩写之类的判断并给予提示，这我没实现，但不难
                input_abbreviation, input_word = abbreviation_separate(input_word)

                # 如果 缩写 和 单词 全对，直接跳出
                if input_word == word and input_abbreviation == abbreviation:
                    promptOpt = 0
                    break
                # 如果拼写正确，但是大小写有误
                elif input_word.lower() == word.lower() and input_abbreviation.lower() == abbreviation.lower():
                    promptOpt = 1

                # 根据错误次数不同，给予不同反馈
                time += 1
                if time == 1:
                    print('不对哦我的宝，你再想想' + promptList[promptOpt])
                elif time == 2:
                    print('再想想，我的儿' + promptList[promptOpt])
                elif time == 3:
                    print('gen ge shanglang lei' + promptList[promptOpt])
        else:   #无缩写，逻辑同上，正确与否仅根据输入单词本身判断
            while time < 3:
                input_word = input('请输入拼写: ')
                if input_word == 'q':
                    order = input('您确定要退出吗[y/n]?')
                    if order == 'y':
                        quit = True
                        break
                elif input_word == '1':
                    print(word)  # 打印答案，不想要可以去掉。你如果想，可以添加将上一个单词重新加入的功能
                    break
                elif input_word == word:
                    promptOpt = 0
                    break
                elif input_word.lower() == word.lower():
                    promptOpt = 1

                time += 1
                if time == 1:
                    print('不对哦我的宝，你再想想' + promptList[promptOpt])
                elif time == 2:
                    print('再想想，我的儿' + promptList[promptOpt])
                elif time == 3:
                    # 留着，可能以后有更详细的功能
                    pass
        # 注意，这里不能忘记对于是否退出的判断，不然输入退出了依然出不去，直到背完所有单词(监狱)
        if not quit:
            # print('time: ', time)
            if time == 0:
                if input_word != '1':   # 这里对于已经会了的单词就没有夸奖了
                    print('不愧是你！', end="\n")

                # 该单词的剩余背诵次数
                val = enemies[pointer][1]

                if val > 0:
                    enemies[pointer][1] -= 1
                    val -= 1    # 注意val要同步减一
                if val == 0:    # 所有次数都完成，从enemies中去除
                    # print(enemies)
                    finshedwords.append(enemies[pointer][0])
                    enemies.pop(pointer)
                    ptr += 1
                    pointer -= 1    # 指针前移
                else:   # 还有几次要背
                    # 精髓之一， 与前面某个单词换位，这里的10代表与距离当前位置10个以内的一个单词换位，可以把10用某个变量代替，然后统一修改
                    # 这样子有可能出现连续让你回答这个单词1,2,3次，你可以根据自己的需要进行修改让连续回答的情况不会出现之类
                    randomSwapInd = random.sample(range(max(0, pointer - 10), pointer + 1), 1)[0]
                    enemies[pointer], enemies[randomSwapInd] = enemies[randomSwapInd], enemies[pointer]

            elif time < 3:
                # 没有一次成功，
                enemies[pointer][1] = 3
                randomSwapInd = random.sample(range(max(0, pointer - 10), pointer + 1), 1)[0]
                enemies[pointer], enemies[randomSwapInd] = enemies[randomSwapInd], enemies[pointer]

            elif time == 3:
                if formerPrompOpt == 1:     # 有缩写
                    print(abbreviation + '(' + word + ')')
                else:   # 无缩写
                    print(word)
                # 换位
                randomSwapInd = random.sample(range(max(0, pointer - 10), pointer + 1), 1)[0]
                enemies[pointer], enemies[randomSwapInd] = enemies[randomSwapInd], enemies[pointer]
            else:   # 暂时没有操作
                pass
        row = len(enemies)
        left = row - 1  #因为excel下标0行是attributes
        if left in seprator:
            print('还有', left, '个单词，坚持一下！')

    if quit:    # 直接退出
        print('休息一下，马上回来！')
    else:   # 背完退出
        print('你很勇嘛少年！快来解锁进阶版吧！23:59分前付款更可享受20折哦！')

    # 对变化了的剩余次数进行更新
    for info in enemies:
        line_num, times_left = info[:2]
        if sheet.cell_value(line_num, colopt) != times_left:
            sheet1.write(line_num, colopt, times_left)
    # 对已完成的单词直接将剩余背诵次数归零
    for finshedwordsInd in finshedwords:
        sheet1.write(finshedwordsInd, colopt, 0)

    # 保存进度并提示
    wt.save(workbookopt)
    print('process saved')

if '__init__==__main___':
    start_from_recode = int(input("请输入是否在现有记录上继续背单词[0:否, 1:是] :"))
    words_killing(start_from_recode)