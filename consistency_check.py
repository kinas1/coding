#kappa系数

import pandas as pd
#把excel传入python对象（列表）
def excel_transfer(path):
    df=pd.read_excel(path)#这个会直接默认读取到这个Excel的第一个表单
    data = df.values#获取整个工作表数据
    datah=[]
    for d in data:
        dh=[]
        for i in d:
            if type(i) == float and 0 < i < 19:
                dh.append(int(i))
            elif type(i) == float:
                dh.append(100)#此处100表示无双重行为编码
            else:
                dh.append(i)
        datah.append(dh)
    return datah

def input_path():
    path1=input("请输入处理文件1的绝对路径(例如C:/Users/test.xlsx):")
    path2=input("请输入处理文件2的绝对路径(例如C:/Users/test.xlsx):")
    data1=excel_transfer(path1)
    data2=excel_transfer(path2)
    return data1,data2,path1,path2

def make_matrix_h(data1,data2,path1):
    matrix_h={}
    for i in range(19,24):
        matrix_h[i]={}
        for j in range(19,24):
            matrix_h[i][j]=0
    if "慕课" in path1:
        for i in range(len(data1)):
            x=data1[i][5]
            y=data2[i][5]
            matrix_h[x][y]+=1
    else:
        for i in range(len(data1)):
            x=data1[i][6]
            y=data2[i][6]
            matrix_h[x][y]+=1
    return matrix_h

#写一个矩阵，看分类
def make_matrix_x(data1,data2,path1):
    d1=[]
    d2=[]#合并教学行为为一列
    if "慕课" in path1:
        for d in data1:
            d1.append(d[4])
        for d in data2:
            d2.append(d[4])
    else:        
        for d in data1:
            d1.append(d[4]*1000+d[5])
        for d in data2:
            d2.append(d[4]*1000+d[5])
            
    d=[]#统计出现的教学行为分类    
    for i in range(len(d1)):
        if d1[i] not in d:
            d.append(d1[i])
        if d2[i] not in d:
            d.append(d2[i]) 
    #构建矩阵
    matrix_x={}
    for i in d:
        matrix_x[i]={}
        for j in d:
            matrix_x[i][j]=0
    for i in range(len(d1)):
        x=d1[i]
        y=d2[i]
        matrix_x[x][y]+=1
    #print_matrix(matrix_d)
    return matrix_x

def check_output(matrix_d):
    po=pe=fz_o=fz_e=sum=0
    sum_column=[]
    sum_row=[]
    keys=[]
    for k in matrix_d:
        keys.append(k)
    for i in keys:
        column=0
        row=0
        for j in keys:
            row+=matrix_d[i][j]
            column+=matrix_d[j][i]
            sum+=matrix_d[i][j]
            if i==j:
                fz_o+=matrix_d[i][j]           
        sum_column.append(column)
        sum_row.append(row)
    
    #print(sum_column,'\n',sum_row,'\n',sum)
    po=fz_o/sum
    for i in range(len(sum_column)):
        fz_e+=sum_column[i]*sum_row[i]
    
    pe=fz_e/(sum*sum)
    #print(fz_e,po,pe)
    kappa=(po-pe)/(1-pe)
    return kappa

def print_out():
    data1,data2,path1,path2=input_path()
    matrix_x=make_matrix_x(data1,data2,path1)
    matrix_h=make_matrix_h(data1,data2,path1)
    kappa_x=check_output(matrix_x)
    kappa_h=check_output(matrix_h)
    print("进行一致性检验的两个文件的路径分别为：")
    print("\t",path1)
    print("\t",path2)
    print(path1.split('/')[6].split('.')[0],"教学行为标注一致性检验的结果为：{:.2%}".format(kappa_x))
    print(path1.split('/')[6].split('.')[0],"教学环节标注一致性检验的结果为：{:.2%}".format(kappa_h))

print_out()
