'''
子函数（1）——读取某一文件夹目录下，及其子目录下所有文件名字/仅适用于最大二级目录
'''

import os
import xlwt
import math
import xlrd
from xlutils.copy import copy

allFileNum=0
subDirNameFront='None'
subDirName='None'
row_2=0

def printPath(level,path,inputSite,inputOption):
	global allFileNum
	global subDirNameFront
	global subDirName
	global row_2
	#打印一个目录下的所有文件和文件夹
	#所有文件夹，第一个字段是次目录的级别
	dirList=[]
	#所有文件
	fileList=[]
	#返回一个列表，其中包含在目录条目的名称
	files=os.listdir(path)#【核心】——————os.listdir(path)
	#先添加目录级别
	dirList.append(str(level))
	for f in files:
		#如果是文件夹dirList.append
		if(os.path.isdir(path+'\\'+f)):
			#排除隐藏文件夹。因为隐藏文件夹过多
			if(f[0]=='.'):
				pass
			else:
				#添加非隐藏文件夹
				dirList.append(f)
				#如果是文件fileList.append
		if(os.path.isfile(path+'\\'+f)):
			#排除隐藏文件
			if (str(f).split('.')[1]=='000'):
				fileList.append(path+'\\'+f)

	#遍历所有文件夹，得到dirList中的文件名
	i_dl=0
	for dl in dirList:	
		 #当一个标志使用，文件夹列表第一个级别不打印
		if(i_dl==0):
			i_dl=i_dl+1
		else:
			#打印至控制台，不是第一个的目录
			#print('-'*(int(dirList[0])),dl)
			#打印目录下的所有文件夹和文件，目录级别+1
			printPath((int(dirList[0])+1),path+'\\'+dl,inputSite,inputOption)
				
	#将所有文件中的数据都汇总到一个excel文件中
	term=1000
	data=[[0 for i in range(12)] for j in range(term)]
	specificDate=[0 for j in range(term)]
	row=0

	i=0
	for f1 in fileList:
		#print(f1)
		#计算总文件个数
		allFileNum=allFileNum+1
		#得到当前文件夹名字（1）
		subDirName=path.split('\\')[4]#根据数据所在位置变更(3 or 4 or 5)！
		#读取文件内容(针对两种读取数据方式)！2014 and 2017 belong to the first,others are belong to the second
		try:
			fi=open(f1)
			lines=fi.readlines()
			
		except:
			fi=open(f1,'r',encoding='UTF-8')#特殊文件添加rb读取
			lines=fi.readlines()
			
		finally:
			#fi=open(f1,'r',encoding='UTF-8')#特殊文件添加rb读取
			#lines=fi.readlines()
			#如果try和except都错误
			listData=[]

			#根据地点记录帅选数据
			#if (inputOption=='no'):
			try:
				for line in lines:
					splitStr=list(str(line).strip('\n').split('\t'))
					if(splitStr[0].split(' ')[0]==str(inputSite)):
							data[row][0]=int(splitStr[0].split(' ')[0])
							data[row][1]=float(splitStr[0].split(' ')[1])
							data[row][2]=float(splitStr[0].split(' ')[2])
							for i_1 in range(3,12): 
								if i_1==7:
									data[row][i_1]=float(splitStr[i_1-2])
								if i_1!=7:
									data[row][i_1]=int(splitStr[i_1-2])	
			except:
				#2016111501数据格式是不同的！
				for line in lines:
					splitStr=list(str(line).strip('\n').split('\t'))
					if(splitStr[0].split(' ')[0]==str(inputSite)):
						for i_2	in range(12):
							if i_2==9 or i_2==3 or i_2==4:
								data[row][i_2]=float(splitStr[i_2-2])
							else:
								#print('i_2 is:',i_2)
								#print(splitStr)
								#print(splitStr[0],splitStr[1],splitStr[2],splitStr[3],splitStr[4],splitStr[5])
								data[row][i_2]=int(splitStr[i_2-2]) 			
			#else:
			#	print('in the second try,没有读取文件,请查看:'+f1)
			#	break
				
			finally:						
				row=row+1			
	
				a=(f1.split('\\')[4]).split('.')[0]
				specificDate[i]=a
				i=i+1
				fi.close	
			
				if (f1==fileList[-1]):
					#创建表格+写入
					workbook=xlwt.Workbook()#创建excel
					sheet1=workbook.add_sheet('sheet1',cell_overwrite_ok=True)#创建sheet
						
					row_1=0
					for line in data:
						col=1
						#赋值内容
						for word in line:
							if col==1:
								sheet1.write(row_1,0,int(specificDate[row_1]))
							sheet1.write(row_1,col,word)
							col+=1
						row_1+=1
					#如果没有文件夹，则新建文件夹
					if os.path.exists('F:\\lang\\saving\\'+inputSite)==False:
						os.mkdir('F:\\lang\\saving\\'+inputSite)
					workbook.save('F:\\lang\\saving\\'+inputSite+'\\'+subDirName+'.xls')
				
				
		
	'''
		#根据选择计算平均值
		if(inputOption=='yes'):
			#将该文件中数据 存入data中
			data=[[0 for i in range(10)] for j in range(term)]
			row=0
			for line in lines:
				splitStr=list(str(line).strip('\n').split('\t'))
				data[row]=splitStr
				row=row+1
			fi.close

			sum_1=[0]*9
			rms_1=[0]*9
			number_1=[0]*9
			#求和&均方根
			for i_1 in range(1,len(data)):
				for j_1 in range(9):
					num=float(data[i_1][j_1+1])
					if (num>1000)or(num<0.1):
						continue
					#计算shu
					number_1[j_1]=number_1[j_1]+1						
					sum_1[j_1]=sum_1[j_1]+num
					rms_1[j_1]=rms_1[j_1]+num
	
			#写入Mean_Rms
			#创建表格+写入
			
			#新建excel
			if(allFileNum==1) or(subDirName!=subDirNameFront):
				#另起文件夹 需要从第0行开始写
				if(subDirName!=subDirNameFront):
					row_2=0
				workbook=xlwt.Workbook()#创建excel
				sheet1=workbook.add_sheet('sheet1',cell_overwrite_ok=True)#创建sheet
				#赋值内容Mean&rms_1
				for j_2 in range(9):
					sheet1.write(row_2,0,(f1.split('\\')[4]).split('.')[0])#对应文件名字
					sheet1.write(row_2,j_2+1,sum_1[j_2]/number_1[j_2])#平均值
					sheet1.write(row_2,j_2+1+10,math.sqrt(rms_1[j_2]/number_1[j_2]))#均方根
				workbook.save(subDirName+'Mean_Rms'+'.xls')
			
			#打开已有excel
			if(allFileNum!=1):
				xls=xlrd.open_workbook('F:\\lang\\'+subDirName+'Mean_Rms'+'.xls')
				xlsc=copy(xls)
				shtc=xlsc.get_sheet(0)
				for j_2 in range(9):
					shtc.write(row_2,0,(f1.split('\\')[4]).split('.')[0])#对应文件名字
					shtc.write(row_2,j_2+1,sum_1[j_2]/(len(lines)-1))#平均值
					shtc.write(row_2,j_2+1+10,math.sqrt(rms_1[j_2]/(len(lines)-1)))#均方根
				xlsc.save('F:\\lang\\'+subDirName+'Mean_Rms'+'.xls')
	
		row_2=row_2+1
		#得到前一个文件夹名字（2）
		subDirNameFront=subDirName	
		
'''

#主函数
if __name__=='__main__':
	#inputOption=input('请输入你的选择（"yes"表示只计算平均值,"no"表示只计算汇总一个位置点数据）：')
	inputOption='no'
	inputSite=input('请输入你要查找的位置：')
	#inputYear=input('请输入你要查找的年限(2014-2018)：')
	for i in range(2014,2019):
		inputYear=str(i)
		printPath(1,'F:\\lang\\data\\'+inputYear,inputSite,inputOption)

