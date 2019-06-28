# -*- coding: utf-8 -*-
"""
Created on Wed Jun 26 12:21:08 2019

@author: lenovo
"""

import sqlite3
import os
import pandas as pd
import cx_Oracle
import numpy as np
from openpyxl import load_workbook

conn = sqlite3.connect('univs.db')
cursor = conn.cursor()
def create_table():

    cursor.execute('''CREATE TABLE Univ
	       (ID INT NOT NULL,
	       Name           TEXT    NOT NULL,
	       Year            INT     NOT NULL,
	       Major            TEXT     NOT NULL,
	       HighestScore            INT ,
	       AverageScore            INT ,
	       LowestScore            INT,
	       LowestRank            INT,
	       Description            TEXT);''')
    conn.commit()

def insert_hunan(site):
    univs_list = os.listdir('%s数据'%site)
    for univ in univs_list:
        univ_name = univ.split('.')[0].split('_')[0]
        univ_id = int(univ.split('.')[0].split('_')[1])
        year = int(univ.split('.')[0].split('_')[2])
        print(univ_name, univ_id, year)
        with open('%s数据/'%site + univ, 'r') as file:
            for line in file:
                line = line.replace('-', '-999')
                major = line.split(' ')[0]
                HighestScore = int(line.split(' ')[1])
                AverageScore = int(line.split(' ')[2])
                LowestScore = int(line.split(' ')[3])
                LowestRank = int(line.split(' ')[4])
                Description = line.split(' ')[5]
                # print(major, HighestScore, AverageScore, LowestScore, LowestRank, Description)
                cursor.execute("insert into univ (ID, Name, YEAR, major, HighestScore, AverageScore, LowestScore, LowestRank, Description) values (?,?,?,?,?,?,?,?,?);", (univ_id, univ_name, year, major, HighestScore, AverageScore, LowestScore, LowestRank, Description))
                conn.commit()
    conn.commit()
    conn.close()

def insert_beijing():
    univs_list = os.listdir('北京数据')
    for univ in univs_list:
        univ_name = univ.split('.')[0].split('_')[0]
        univ_id = int(univ.split('.')[0].split('_')[1])
        year = int(univ.split('.')[0].split('_')[2])
        # print(univ_name, univ_id, year)
        with open('北京数据/' + univ, 'r') as file:
            for line in file:
                line = line.replace('-', '-999')
                major = line.split(' ')[0]
                HighestScore = int(line.split(' ')[1])
                AverageScore = int(line.split(' ')[2])
                LowestScore = int(line.split(' ')[3])
                LowestRank = int(line.split(' ')[4])
                Description = line.split(' ')[5]
                # print(major, HighestScore, AverageScore, LowestScore, LowestRank, Description)
                cursor.execute("insert into univ (ID, Name, YEAR, major, HighestScore, AverageScore, LowestScore, LowestRank, Description, site) values (?,?,?,?,?,?,?,?,?, '北京');", (univ_id, univ_name, year, major, HighestScore, AverageScore, LowestScore, LowestRank, Description))
                conn.commit()
    conn.commit()
    conn.close()

def create_excel(site):
    path = '%s高校各专业在湖南历年分数线及省排名（理科）.xlsx'%site
    columes_rename_dict = {'ID': '编号', 'Name': '高校', 'Year': '年份', 'Major': '专业', 'HighestScore': '最高分', 'AverageScore': '平均分', 'LowestScore': '最低分', 'LowestRank': '最低省排名', 'Description': '录取批次', 'site': '所在省市', 'ApproxRank': '省排名（估计）'}
    df = pd.read_sql("select * from univ where site='%s' order by ID, year desc"%site , conn)
    df.columns = ['编号', '高校', '年份', '专业', '最高分', '平均分', '最低分', '最低省排名', '录取批次', '所在省市', '省排名（估计）']
    for col in df.columns:
        df[col][df[col] == -999] = np.nan
    # df.rename(index=str, columns = columes_rename_dict)
    df.to_excel(path)


    book = load_workbook(path)
    writer = pd.ExcelWriter(path, engine = 'openpyxl')
    writer.book = book
    univs_list = os.listdir('%s数据'%site)
    for univ in univs_list:
        univ_name = univ.split('.')[0].split('_')[0]
        univ_id = int(univ.split('.')[0].split('_')[1])
        year = int(univ.split('.')[0].split('_')[2])
        print(univ_name)
        if year==2018:
            print(univ_name, univ_id)
            df = pd.read_sql('select * from univ where id=%d order by ID, year desc' % univ_id, conn)
            for col in df.columns:
                df[col][df[col] == -999] = np.nan
            df.columns = ['编号', '高校', '年份', '专业', '最高分', '平均分', '最低分', '最低省排名', '录取批次', '所在省市', '省排名（估计）']
            # df.rename(index=str, columns=columes_rename_dict)
            df.to_excel(writer, sheet_name=univ_name)
    writer.save()
    writer.close()

def add_approxrank(site='湖南'):
    lookup_dict = {}
    for year in range(2014, 2019):
        lookup_dict[year] = pd.read_excel('%d年理科一份一段表.xlsx' % year)
        # print(lookup_dict[year].dtypes)
    # lookup_2018 = pd.read_excel('2018年理科一份一段表.xlsx')
    df = pd.read_sql("select * from univ where site='%s' order by ID, year desc" % site , conn)
    print(df.dtypes)
    for index, row in df.iterrows():
        print('%d of %d' % (index, len(df)))
        ID = row['ID']
        Name = row['Name']
        Year = row['Year']
        Major = row['Major']
        HighestScore = row['HighestScore']
        AverageScore = row['AverageScore']
        LowestScore = row['LowestScore']
        LowestRank = row['LowestRank']
        if LowestRank==-999:#如果没有排名
            if LowestScore != -999:
                this_rank = lookup_dict[Year].query('档分==%d'%LowestScore)['累计人数2'].values[0]
            elif AverageScore != -999:
                this_rank = lookup_dict[Year].query('档分==%d' % AverageScore)['累计人数2'].values[0]
            elif HighestScore != -999:
                this_rank = lookup_dict[Year].query('档分==%d' % HighestScore)['累计人数2'].values[0]
            else:
                this_rank = -999
            cursor.execute("update univ set ApproxRank=%d where ID=%d and Name='%s' and Year=%d and Major='%s'" % (this_rank, ID, Name, Year, Major))
            conn.commit()



# insert_beijing()
# add_approxrank('北京')
# create_excel('北京')
# insert_hunan('广东')
add_approxrank('广东')
create_excel('广东')
conn.commit()
conn.close()