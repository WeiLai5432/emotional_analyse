#encoding=utf-8
# import jieba
import openpyxl
from openpyxl import load_workbook
# import jieba.posseg as pseg
import json

# jieba.enable_parallel(4) # 
trainfile = load_workbook('trainfile.xlsx')
ws = trainfile.get_sheet_by_name(trainfile.get_sheet_names()[0])
rows = ws.rows
col = ws.columns

newdic = open("newdic.txt",'w')
plms = open("plms.txt",'w')
newmap = {}

for i, row in enumerate(rows):
	if i == 0:
		continue
	word = []

	for j, cell in enumerate(row):
		if j == 0:
			print cell.value
		# if j == 1:
		# 	val = unicode(cell.value)
		# 	words = pseg.cut(val)
		# 	qgword = [x for x, f in words if f == 'a']
			# print json.dumps(qgword, encoding="UTF-8", ensure_ascii=False)
		if j == 2:
			if cell.value != None:
				a = cell.value.split(';')
				for aa in a:
					if aa != '' and aa != 'NULL' and aa not in newmap.keys():
						newdic.write(aa.encode('utf-8') + ' n' + '\n')
						newmap[aa] = 1
		if j == 3:
			if cell.value != None:
				a = cell.value.split(';')
				for aa in a:
					if aa != '' and aa not in word:
						newdic.write(aa.encode('utf-8') + ' a' + '\n')
						word.append(aa)
		if j == 4:
			if cell.value != None:
				a = cell.value.split(';')
				# print a
				for (aa,bb) in zip(a, word):
					if aa != '':
						plms.write(bb.encode('utf-8')+' '+aa.encode('utf-8') + '\n')

# for word, flag in words:
# 	print ('%s %s' % (word,flag))