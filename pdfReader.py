from io import StringIO
import os
import re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from openpyxl import Workbook
from functools import wraps
from threading import Thread
import functools

def timeout(seconds_before_timeout):
	def deco(func):
		@functools.wraps(func)
		def wrapper(*args, **kwargs):
			res = [Exception('function [%s] timeout [%s seconds] exceeded!' % (func.__name__, seconds_before_timeout))]
			def newFunc():
				try:
					res[0] = func(*args, **kwargs)
				except Exception as e:
					res[0] = e
			t = Thread(target=newFunc)
			t.daemon = True
			try:
				t.start()
				t.join(seconds_before_timeout)
			except Exception as e:
				print('error starting thread')
				raise e
			ret = res[0]
			if isinstance(ret, BaseException):
				print("moving on")
			return ret
		return wrapper
	return deco

@timeout(30)
def convert(fname, pages=None):
	if not pages:
		pagenums = set()
	else:
		pagenums = set(pages)
	output = StringIO()
	manager = PDFResourceManager()
	converter = TextConverter(manager, output, laparams=LAParams())
	interpreter = PDFPageInterpreter(manager, converter)

	infile = open(fname, 'rb')
	for page in PDFPage.get_pages(infile, pagenums, check_extractable=False):
		interpreter.process_page(page)
		if "Keywords" in output.getvalue():
			text = output.getvalue()
			m = re.search('(?<=Keywords)(.*)(?=\n)',text)
			if not re.search('[a-zA-Z]', m.group(0)):
				m = re.search('(?s)(?<=Keywords)(.*?)(?:(?:\r*\n){2})',text)
				m = re.sub("\n",",",m)
			kw = m.group(0)
			print(kw)
			return kw
		if "Key words" in output.getvalue():
			text = output.getvalue()
			m = re.search('(?<=Key words)(.*)(?=\n)',text)
			kw = m.group(0)
			print(kw)
			return kw
	infile.close()
	converter.close()
	output.close
	


if __name__ == '__main__':
	wb = Workbook()
	ws = wb.active
	fileList = []
	#recursively hit all files
	for (dirname, dirs, files) in os.walk('.'):
		for filename in files:
			thefile = os.path.join(dirname,filename)
			print(thefile)
			if filename.endswith('.pdf') :
				kw = convert(thefile)
				if not isinstance(kw, Exception):
					if kw is not None:
						ws.append([thefile,re.sub('[^A-Za-z0-9]+', ',', kw)])
					else:
						fileList.append(thefile)
			else:
				fileList.append(thefile)
	wb.save("sample.xlsx")
	print(fileList)
	with open('your_file.txt', 'w') as f:
		for item in fileList:
			f.write("%s\n" % item)
	breakpoint()