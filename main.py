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

def read_pdf(pdf_url, sheet, count):
	try:
		r = requests.get(pdf_url)
		if r.status_code != 200:
			return False
		
		pdf_file_path = os.getcwd() + r"\temp.pdf";
		with open(pdf_file_path, 'wb') as f:
			f.write(r.content)
		#pdf_file_path = r"C:\Work\python\CAPCL\RLCA_LU12_08-28-20_Redacted.pdf"
		fp = open(pdf_file_path,'rb')
		# 创建一个与文档关联的解释器
		parser = PDFParser(fp)
		# PDF文档对象
		doc = PDFDocument()

		# 链接解释器和文档对象
		parser.set_document(doc)

		doc.set_parser(parser)
		# 初始化文档
		doc.initialize()

		# 检测文档是否提供txt转换，不提供就忽略
		if not doc.is_extractable:
			return False

		# 创建PDF资源管理器
		resource = PDFResourceManager()
		# 参数分析器
		laparam = LAParams()
		# 创建聚合器
		device = PDFPageAggregator(resource, laparams=laparam)
		# 页面解释器
		interpreter = PDFPageInterpreter(resource, device)

		REC_flag = False
		REP_flag = False
		reporting_flag = True
		Recordkeeping_V_flag = True
		Reporting_V_flag = True
		REC_cnt = 0
		REP_cnt = 0
		data_index = 0
		# 使用文档对象得到页面内容
		for page in doc.get_pages():
			# 使用页面解释器读取
			interpreter.process_page(page)
			# 使用聚合器获得内容
			layout = device.get_result()
			for out in layout:
				if hasattr(out, "get_text"):
					data_index = data_index + 1
					if out.get_text().strip() == "":
						continue
					
					if "LM Number:" in out.get_text():
						str_strat = out.get_text().find("LM Number:") + len("LM Number:")
						sheet.write(count, 3, (out.get_text())[str_strat:].strip())
					elif "LMNumber:" in out.get_text().strip().replace(" ", ""):
						str_strat = out.get_text().strip().replace(" ", "").find("LMNumber:") + len("LMNumber:")
						sheet.write(count, 3, (out.get_text().strip().replace(" ", ""))[str_strat:].strip())
					elif "LMNmnber:" in out.get_text().strip().replace(" ", ""):
						str_strat = out.get_text().strip().replace(" ", "").find("LMNmnber:") + len("LMNmnber:")
						sheet.write(count, 3, (out.get_text().strip().replace(" ", ""))[str_strat:].strip())
					elif "LMNlllllber:" in out.get_text().strip().replace(" ", ""):
						str_strat = out.get_text().strip().replace(" ", "").find("LMNlllllber:") + len("LMNlllllber:")
						sheet.write(count, 3, (out.get_text().strip().replace(" ", ""))[str_strat:].strip())

					if "the following recordkeeping violations:" in out.get_text() \
						or "the following recordkeeping violation:" in out.get_text() \
						or "thefollowingrecordkeepingviolation:" in out.get_text().strip().replace(" ", "").replace("\n", ""):
						sheet.write(count, 4, out.get_text())

					if Recordkeeping_V_flag and ("Recordkeeping Violations" in out.get_text() \
						or "Recordkeeping Violation" in out.get_text()
						or "RecordkeepingViolation" == out.get_text().strip().replace(" ", "")
						or "RecordkeepingViolations" == out.get_text().strip().replace(" ", "")):
						Recordkeeping_V_flag = False
						REC_flag = True
						sheet.write(count, 5, "1")

					if REC_flag:
						if re.match("^[0-9].*", out.get_text()):
							REC_cnt = REC_cnt + 1

					if reporting_flag and "for the fiscal year ended" in out.get_text():
						reporting_flag = False
						str_strat = out.get_text().rfind(".") + len(".")
						sheet.write(count, 7, (out.get_text())[str_strat:].strip())

					if Reporting_V_flag and "Reporting Violations" in out.get_text():
						Reporting_V_flag = False
						REC_flag = False
						REP_flag = True
						sheet.write(count, 8, "1")

					if data_index == 2:
						sheet.write(count, 10, out.get_text())
					
					if data_index == 4:
						sheet.write(count, 11, out.get_text())

					if REP_flag:
						if re.match("^[0-9].*", out.get_text()):
							REP_cnt = REP_cnt + 1

					if "OtherIssues" == out.get_text().strip().replace(" ", "") \
						or "OtherViolation" == out.get_text().strip().replace(" ", ""):
						REC_flag = False
						REP_flag = False

		sheet.write(count, 6, str(REC_cnt))
		sheet.write(count, 9, str(REP_cnt))
		fp.close()
		if(os.path.exists(pdf_file_path)):
			os.remove(pdf_file_path)
	except:
		print("pdf 解析失败", flush = True)
		return False
	return True

if __name__ == "__main__":
	#sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
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
	count=1
	workbook = xlwt.Workbook()
	sheet = workbook.add_sheet("Sheet Name1")
	sheet.write(0, 0, "Union Name")
	sheet.write(0, 1, "Affiliate")
	sheet.write(0, 2, "Date")
	sheet.write(0, 3, "LM_Number")
	sheet.write(0, 4, "recordkeeping")
	sheet.write(0, 5, "recordkeeping_violations")
	sheet.write(0, 6, "number_recordkeeping_v")
	sheet.write(0, 7, "reporting")
	sheet.write(0, 8, "reporting_violations")
	sheet.write(0, 9, "number_recordkeeping_v")
	sheet.write(0, 10, "office_zip")
	sheet.write(0, 11, "union_zip")

	try:
		for year in range(2016,last_year):
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
				CAPDataArray = j.text.split("\n")
				CAPDataArray = [x for x in CAPDataArray if x!='']
				CAPDataArray = [x for x in CAPDataArray if x!='\xa0']

				print("Union Name:" + str(CAPDataArray[0]), flush = True)
				print("Affiliate:" + CAPDataArray[1], flush = True)
				print("Date:" + CAPDataArray[2], flush = True)
				print("--------------------------")
				sheet.write(count,0, CAPDataArray[0]) # row, column, value
				sheet.write(count,1, CAPDataArray[1])
				sheet.write(count,2, CAPDataArray[2])

				print(CAPDataArray, flush = True)
				pdf_url = "https://www.dol.gov"
				if len(CAPDataArray) == 4:
					if year == 2016 and CAPDataArray[0] == "United Nurses and Allied Professionals":
						pdf_url = pdf_url + (j.contents)[9].select("a")[0]['href']
					else:	
						pdf_url = pdf_url + (j.contents)[7].select("a")[0]['href']
				else:
					if "HTML" in CAPDataArray or CAPDataArray[3] == "-":
						pdf_url = pdf_url + (j.contents)[9].select("a")[0]['href']
					else:
						pdf_url = pdf_url + (j.contents)[7].select("a")[0]['href']
				#print("pdf_url:" + pdf_url, flush = True)
				ret = read_pdf(pdf_url, sheet, count)
				count = count + 1
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
			obj_f.write(CAPDataArray[1] + "," + CAPDataArray[2] + "," + CAPDataArray[3])

	# 执行完成后，删除文件
	if(os.path.exists('./the_last_dance.txt')):
		os.remove('./the_last_dance.txt')

	print("完成",flush = True)
