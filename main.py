#!/usr/bin/python
# -*- coding:utf-8 -*-

import numpy, xlrd, xlwt

# 从文本中读取边缘线(非海岸线)经纬度
def getDots():
    # 经度
    longitude = []
    # 纬度
    latitude = []
    with open('./edge-line.txt') as file_to_read:
        while True:
            line = file_to_read.readline()
            if not line:
                break
            [la, lo] = line.split(',')
            longitude.append(float(lo.replace('\n','')))
            latitude.append(float(la.replace('\n','')))
        return [longitude, latitude]

# 拟合多项多次方程
def polyfit(x, y, degree):
    results = {}
    coeffs = numpy.polyfit(x, y, degree)
    results['polynomial'] = coeffs.tolist()

    p = numpy.poly1d(coeffs)
    yhat = p(x)   
    ybar = numpy.sum(y)/len(y)  
    ssreg = numpy.sum((yhat-ybar)**2)
    sstot = numpy.sum((y - ybar)**2) 
    results['determination'] = ssreg / sstot 
    return results

# 判断某个坐标是否在区域内
def isInner(longitude, latitude, polynomial):
    if longitude == 0 or latitude == 0:
        return False
    n = len(polynomial)
    index = 0
    sum = 0
    while n > 0:
        sum += polynomial[n-1] * longitude**index
        n -= 1
        index += 1
    return sum >= latitude

# 读取excel数据
def readExcel():
    print 'reading ...'
    book = xlrd.open_workbook("data.xls")
    print 'excel read success'
    return book.sheet_by_index(0)

# excel写数据
def writeExcel(sheet, polynomial):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('茕茕孑立沆瀣一气')

    nrows = sheet.nrows
    ncols = sheet.ncols

    for row in range(0, 200):
        for col in range(0, ncols):
            ws.write(row, col, sheet.cell_value(rowx=row, colx=col))

        if row != 0:
            longitude = sheet.cell_value(row, 6)
            latitude = sheet.cell_value(row, 7)
            status = isInner(float(longitude), float(latitude), polynomial)
            ws.write(row, ncols, status)
            if status:
                print row,sheet.cell_value(rowx=row, colx=1),sheet.cell_value(rowx=row, colx=2),latitude,longitude
        else:
            ws.write(row, ncols, '是否在区域内')

    wb.save('out/example.xls')
    

if __name__ == '__main__':
    [longitudes, latitudes] = getDots()
    z1 = polyfit(longitudes, latitudes, 3)
    sheet = readExcel()
    writeExcel(sheet, z1['polynomial'])