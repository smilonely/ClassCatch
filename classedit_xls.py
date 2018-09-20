#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import time  #时间戳相关
import datetime  #Datetime相关
import random  #生成随机数
import string  #生成随机UID相关
import re  #多字符分割字符相关
import codecs  #使用utf-8编码
import xlrd  #编辑xls

DELIMITER_0 = "◇|\n"  #用于各个部分的分割
DELIMITER_1 = "◇"  #用来计数
DELIMITER_2 = "("  #分割周数与课程节次
DELIMITER_3 = ","  #分割课程节次（然并卵），主要是用于分析单个Slice有几个周区间
DELIMITER_4 = "-|,"  #分割周区间
#规定分割课表的几个常量，可以根据课表格式更改
#示例：「 大学英语1◇1-16(1,2节)◇教1-137◇教师名」

DELIMITER_5 = "("
DELIMITER_6 = "："
#分割班级名
#示例：「班级名称：电竞1801(38人)    班级代码：201810101」

DELIMITER_7 = "("
DELIMITER_8 = ")"
DELIMITER_9 = "-"
#分割万恶的体育课
#示例：「 大学体育1◇2节/周(1-16)[7-8节]」

UID = "1234567890abcdefghijklmnopqrstuvwxyz\
ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*"
#UID备选列表

def begin_time():
	ver = 0
	while ver == 0:
	#循环检查输入是否有误，ver为1为有误，1为正确
		ver = 1

		try:
			time = input ()
			date_time = datetime.datetime.strptime(time, "%Y-%m-%d-%H-%M")
			#输入日期并转换成datetime

		except:
			print ("\r\n输入有误，请重新输入：\r\n")
			ver = 0
	return date_time
#获取时间并转化成datetime

def time_verification(date_time, num):
	while date_time - start_list[num] < datetime.timedelta(minutes = 1):
		print ("\r\n输入时间早于课程开始时间，请检查并重新输入：\r\n")
		
		time = input ()
		date_time = datetime.datetime.strptime(time, "%Y-%m-%d-%H-%M")

	return date_time
#验证时间早晚

start_list = [i for i in range(0, 13)]
#课程开始时间list
end_list = [i for i in range(0, 13)]
#课程结束时间list

print ("\r\n请输入第一周星期一上午第一节课上课日期和时间\r\n")
print ("（无论第一节有没有课），\r\n")
print ("格式为「年-月-日-时-分」（24小时制，不足两位请补0），\r\n")
print ("例如：2001-01-01-09-59。\r\n")
print ("请输入：\r\n")
start_list[1] = begin_time()
#start_list[1] = datetime.datetime.strptime("2018-09-10-08-30", "%Y-%m-%d-%H-%M")

print ("\r\n好的！下面再输入星期一下午第一节课的上课日期和时间，\r\n")
print ("也就是第五节课，\r\n")
print ("格式相同。\r\n")
start_list[5] = time_verification(begin_time(), 1)
#start_list[5] = datetime.datetime.strptime("2018-09-10-13-40", "%Y-%m-%d-%H-%M")

print ("\r\n最后，再输入星期一晚上第一节课的上课日期和时间，\r\n")
print ("我没记错的话，应该就是第九节课，\r\n")
print ("格式依然相同。\r\n")
start_list[9] = time_verification(begin_time(), 5)
#start_list[9] = datetime.datetime.strptime("2018-09-10-18-20", "%Y-%m-%d-%H-%M")

print("\r\n好的，请稍等……\r\n")


def plus_60(time):
	time_after = time + datetime.timedelta(minutes = 60)
	return time_after

def plus_50(time):
	time_after = time + datetime.timedelta(minutes = 50)
	return time_after

def plus_45(time):
	time_after = time + datetime.timedelta(minutes = 45)
	return time_after

def plus_day(time, day):
	time_after = time + datetime.timedelta(days = day)
	return time_after

def plus_week(time, week):
	time_after = time + datetime.timedelta(weeks = week)
	return time_after


for i in range(0, 3):
	start_list[2+4*i] = plus_50(start_list[1+4*i])
	start_list[3+4*i] = plus_60(start_list[2+4*i])
	start_list[4+4*i] = plus_50(start_list[3+4*i])
	for l in range(1,5):
		end_list[l+4*i] = plus_45(start_list[l+4*i])
#推算每节课的时间，注意这个list储存的是小课时间


def uid_born():
        uid_list = []
        for i in range(0,20):
                uid_list.append(random.choice(UID))
                uid = "".join(uid_list)
        return uid
#随机生成UID


def classedit_xls(data_slice_1, y, x):		
	global write_text
	#引入全局变量 write_text

	name = data_slice_1[0]
	name = " ".join(name.split())  #去掉课程名字前的空格
	week = week_list[x-1]		
	
	ver_2 = 1
	ver_3 = 1
	#两个验证变量，ver_2为地点，ver_3为描述（教师）
	try:
		location = data_slice_1[2]
	except:
		ver_2 = 0
	try:
		description = "教师：" + data_slice_1[3]
	except:
		ver_3 = 0
	#验证是否有地点和教师
	#给体育课留的，体育课已经单独分出了函数，现在看了并没有什么卵用（摊手）

	data_slice_2 = data_slice_1[1].split(DELIMITER_2)
		
	loop = data_slice_1[1].count(DELIMITER_3)
	for i in range(0, loop):
		data_slice_3 = re.split(DELIMITER_4, data_slice_2[0])
			
		start_week = int(data_slice_3[0+2*i]) - 1
		duration_week = int(data_slice_3[1+2*i]) - start_week
		#计算每节课起始周与持续周数
		
		start_lesson = plus_week(start_list[y*2-5], start_week)  #计算每节大课开始时间
		start_lesson = plus_day(start_lesson, x-2)  #计算每节课在一周中的偏移
		end_lesson = plus_50(plus_45(start_lesson))  #计算每节大课下课时间
		start_lesson = start_lesson.strftime("%Y%m%dT%H%M%S")
		end_lesson = end_lesson.strftime("%Y%m%dT%H%M%S")  #格式化时间

		write_text.append("BEGIN:VEVENT")
		write_text.append("UID:" + uid_born() + "@github.com/smilonely")
		write_text.append("DTSTART;TZID=Asia/Shanghai:" + start_lesson)
		write_text.append("DTEND;TZID=Asia/Shanghai:" + end_lesson)
		write_text.append("RRULE:FREQ=WEEKLY;WKST=SU;COUNT=" + \
			str(duration_week) + ";BYDAY=" + week)
		if ver_2 == 1:
			write_text.append("LOCATION:" + location)
		if ver_3 == 1:
			write_text.append("DESCRIPTION:" + description)
		write_text.append("STATUS:CONFIRMED")
		write_text.append("SUMMARY:" + name)
		write_text.append("TRANSP:OPAQUE")
		write_text.append("END:VEVENT\r\n")
#添加一个EVENT的内容进list write_text，但是不合并


def class_pe_edit (data_slice_1, y, x):
	global write_text

	name = data_slice_1[0]
	name = " ".join(name.split())  #去掉课程名字前的空格
	week = week_list[x-1]

	data_slice_2 = data_slice_1[1].split(DELIMITER_8)
	data_slice_2 = data_slice_2[0].split(DELIMITER_7)
	data_slice_2 = data_slice_2[1].split(DELIMITER_9)
	#最后分割出体育课的开始周与结束周
	#实际上并没有什么卵用，看起来所有体育课都是1-16周
			
	start_week = int(data_slice_2[0]) - 1
	duration_week = int(data_slice_2[1]) - start_week
		
	start_lesson = plus_week(start_list[y*2-5], start_week)
	start_lesson = plus_day(start_lesson, x-2)
	end_lesson = plus_50(plus_45(start_lesson))
	start_lesson = start_lesson.strftime("%Y%m%dT%H%M%S")
	end_lesson = end_lesson.strftime("%Y%m%dT%H%M%S")

	write_text.append("BEGIN:VEVENT")
	write_text.append("UID:" + uid_born() + "@github.com/smilonely")
	write_text.append("DTSTART;TZID=Asia/Shanghai:" + start_lesson)
	write_text.append("DTEND;TZID=Asia/Shanghai:" + end_lesson)
	write_text.append("RRULE:FREQ=WEEKLY;WKST=SU;COUNT=" + \
		str(duration_week) + ";BYDAY=" + week)
	write_text.append("STATUS:CONFIRMED")
	write_text.append("SUMMARY:" + name)
	write_text.append("TRANSP:OPAQUE")
	write_text.append("END:VEVENT\r\n")
#抓体育课


data_raw = xlrd.open_workbook("1.xls")
table = data_raw.sheets()[0]
class_name = table.cell(1, 0).value
class_name = class_name.split(DELIMITER_5)
class_name = class_name[0].split(DELIMITER_6)
class_name = class_name[1]
semester = table.cell(1, 7).value
#找到班级和学期名字

new_ics = codecs.open(class_name + "课程表.ics", "w", "utf-8")
new_ics.write(
"BEGIN:VCALENDAR\r\n"
"VERSION:2.0\r\n"
"PRODID:-//DANGEL//Class Schedule//CN\r\n"
"NAME:" + semester + "课程表\r\n"
"X-WR-CALNAME:" + semester + "课程表\r\n"
"DESCRIPTION:" + class_name + "班" + semester + "课程表\r\n"
"X-WR-CALDESC:" + class_name + "班" + semester + "课程表\r\n"
"CALSCALE:GREGORIAN\r\n"
"METHOD:PUBLISH\r\n"
"X-WR-TIMEZONE:Asia/Shanghai\r\n"
"BEGIN:VTIMEZONE\r\n"
"TZID:Asia/Shanghai\r\n"
"X-LIC-LOCATION:Asia/Shanghai\r\n"
"BEGIN:STANDARD\r\n"
"TZOFFSETFROM:+0800\r\n"
"TZOFFSETTO:+0800\r\n"
"TZNAME:CST\r\n"
"DTSTART:19700101T000000\r\n"
"END:STANDARD\r\n"
"END:VTIMEZONE\r\n\r\n")
#写入各项信息

write_text = []
week_list = [0,"MO", "TU", "WE", "TH", "FR", "SA", "SU"]

for x in range(2, 9):
	for y in range(3, 9):
		data = table.cell(y, x).value
		number = data.count(DELIMITER_1)
		#数一数分隔符
		data_slice = re.split(DELIMITER_0, data)
		#将一个单元格的内容连同换行符分割
		
		if number == 0:
			continue

		if number > 3:
			for k in range(0, number//3):
				data_slice_0 = data_slice[0+4*k:4+4*k]
				classedit_xls(data_slice_0, y, x)

			write_done = "\r\n".join(write_text)
			new_ics.write(write_done)
			#将多个EVENT合并并写入

		if number == 1:		
			class_pe_edit(data_slice, y, x)
			write_done = "\r\n".join(write_text)
			new_ics.write(write_done)

		if number == 3:
			classedit_xls(data_slice, y, x)
			write_done = "\r\n".join(write_text)
			new_ics.write(write_done)
			#直接将一个EVENT合并并写入

		write_text = [""]
		#初始化变量以重复写入

#三重循环，最为致命（雾）。三个循环来遍历单元格

new_ics.write(u"END:VCALENDAR")
new_ics.close()



print ("\r\n完成了！\r\n")
input()
