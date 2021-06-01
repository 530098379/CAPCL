#!/usr/bin/env python3

import requests
from bs4 import BeautifulSoup
from pathlib import Path
import io
import sys
import re
import xlwt
import xlrd
import os
import time
import datetime
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter

if __name__ == "__main__":
	sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
	print("开始", flush = True)
	add_flag = False #字符串拼接flag
	question_text = "" #输出字符串
	last_num = 0 #最后一次的工会编号
	last_year = datetime.datetime.now().year # 最后一次的年份，默认值为本年
	file_year = ""
	# 文件名
	excel_file_name = os.getcwd() + "\\result_" + \
		datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".xls"

	# 获取最后一次的工会编号以及年份
	#my_file = Path("./the_last_dance.txt")
	#if my_file.is_file():
		#with open('./the_last_dance.txt', 'r') as f:
			#last_data = (f.readline()).split(",")
			#last_num = int(last_data[0])
			#last_year = int(last_data[1])

	# 做成Excel文件
	out_flag = False
	count=0
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet("Sheet Name1")

	try:
		for year in range(last_year - 1,last_year):
			if out_flag:
				break
			# 获取cookie
			url_cok = "https://www.dol.gov/agencies/olms/audits/" + str(year)
			r_cok = requests.get(url_cok)
			cookie_jar = r_cok.cookies

			# 再次封装，获取具体标签内的内容
			result_union = r_cok.text
			bs_union = BeautifulSoup(result_union,"html.parser")

			# 获取已爬取内容中的Fiscal Year行的链接
			data_union = bs_union.select("tbody tr")

			# 循环打印输出
			for j in data_union:
				pdf_url = "https://www.dol.gov" + (j.contents)[7].select("a")[0]['href']
				CAPDataArray = j.text.split("\n")

				print("Union Name:" + str(CAPDataArray[1]), flush = True)
				print("Affiliate:" + CAPDataArray[2], flush = True)
				print("Date:" + CAPDataArray[3], flush = True)
				print("--------------------------")
				sheet.write(count,0, CAPDataArray[1]) # row, column, value
				sheet.write(count,1, CAPDataArray[2])
				sheet.write(count,2, CAPDataArray[3])

				fp = open('C:\Work\python\CAPCL\ATU_LU618_01-23-20_Redacted.pdf','rb')

				# 创建一个与文档关联的解释器
				parser = PDFParser(fp)
				# PDF文档对象
				doc = PDFDocument()
				# 链接解释器和文档对象
				parser.set_document(doc)
				doc.set_parser(parser)
				# 初始化文档
				doc.initialize("")
				# 创建PDF资源管理器
				resource = PDFResourceManager()
				# 参数分析器
				laparam = LAParams()
				# 创建聚合器
				device = PDFPageAggregator(resource, laparams=laparam)
				# 页面解释器
				interpreter = PDFPageInterpreter(resource, device)

				# 使用文档对象得到页面内容
				for page in doc.get_pages():
					# 使用页面解释器读取
					interpreter.process_page(page)
					# 使用聚合器获得内容
					layout = device.get_result()
					for out in layout:
						if hasattr(out, "get_text"):
							#print(out.get_text(), flush = True)
							
							if "LM Number:" in out.get_text():
								str_strat = out.get_text().find("LM Number:") + len("LM Number:")
								sheet.write(count, 3, (out.get_text())[str_strat:].strip())
								#out_flag = True
								#break

							if "the following recordkeeping violations:" in out.get_text():
								sheet.write(count, 4, out.get_text())

							if "Recordkeeping Violations" in out.get_text():
								sheet.write(count, 5, "1")

							if "fiscal year ended" in out.get_text():
								sheet.write(count, 6, out.get_text())

							if "Reporting Violations" in out.get_text():
								sheet.write(count, 7, "1")

				count = count + 1
				# 延迟2秒，防止访问太快
				time.sleep(2)
			out_flag = True
			# 输出结果到Excel
			workbook.save(excel_file_name)

			# 释放变量内存
			del r_cok
			del url_cok
			del result_union
			del bs_union
			del data_union

	finally:
		# 中断或者异常，记录最后的工会编码以及年份
		with open('./the_last_dance.txt', 'w') as obj_f:
			obj_f.write(str(1) + "," + year)

	# 执行完成后，删除文件
	if(os.path.exists('./the_last_dance.txt')):
		os.remove('./the_last_dance.txt')

	print("完成",flush = True)
