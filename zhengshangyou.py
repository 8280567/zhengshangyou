from aip import AipOcr
from PIL import Image, ImageGrab
import jieba
import xlrd
import re

""" 调用通用文字识别, 图片参数为本地图片 """  
APP_ID = '百度通用识别key'
API_KEY = '百度通用识别key'
SECRET_KEY = '百度通用识别key'

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

def remove(text):
    remove_chars = '[0-9’!"#$%&\'()*+,-./:;<=>?@，。?★、…【】《》？“”‘’！[\\]^_`{|}~]+'
    return re.sub(remove_chars, '', text)


def main(a):
	
	#保存截图
	IMAGE_PATH='1.png'
	image = ImageGrab.grab(bbox=(23,310,333,568))
	image.save(IMAGE_PATH)
	
	#识别
	client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
	image = get_file_content('1.png')
	word=client.basicGeneral(image)
	#print(word)
	result=word['words_result']
	#print(result[0])
	##qiege
	cutsearch=[]
	for i in result:
	    for key,value in i.items():
	        #print(key,value)#好用！！！
	        qiepian=jieba.cut(remove(value))
	        cutque=' '.join(qiepian)
	        cutsearch.append(cutque.split(" "))
	#print(cutsearch[1][1])
	#以下是匹配寻找ok
	workbook=xlrd.open_workbook('题库.xlsx')
	sheet=workbook.sheets()[0]
	#第B列
	lie=sheet.col_values(0)
	#print(lie)
	#print(type(useque[0]))
	b=-1
	x=''
	y=''
	z=''
	c=''
	n=''
	try:
		x=cutsearch[0][0]
		y=cutsearch[0][1]
		z=cutsearch[0][2]
		c=cutsearch[1][1]
		n=cutsearch[2][1]
		print('关键字：'+x,y,z,c,n) 
	except:
		print('')#pleas more key
	if x=='':
		pass
	else:
		for i in lie:
			b=b+1
			if re.match('.*?'+x   +   '.*?'+y +   '.*?'+z  +   '.*?'+c  +  '.*?'+n  +'.+',i)!=None:
				print(b)
				dan=sheet.cell_value(b,5)#行，列
				print('**********************'+dan)

	
if __name__ == '__main__':
	for a in range(1,59):
		print('第'+str(a)+'局')
		start=input('任意键')
		main(a)

	





