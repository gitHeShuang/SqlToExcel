import pymysql
import xlwt
import operator
'sql导出excel方法'
def toExcel(host,port,user,passwd,db,charset,sql,sheet_name,out_path):
    db = pymysql.connect(
        host=host,
        port=port,
        user=user,
        passwd=passwd,
        db=db,
        charset=charset
    )
    cursor = db.cursor()
    count = cursor.execute(sql)
    print("查询出" + str(count) + "条记录")

    # 来重置游标的位置
    cursor.scroll(0, mode='absolute')
    # 搜取所有结果
    results = cursor.fetchall()

    # 获取MYSQL里面的数据字段名称
    fields = cursor.description
    workbook = xlwt.Workbook()  # workbook是sheet赖以生存的载体。
    sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)

    # 写上字段信息
    for field in range(0, len(fields)):
        sheet.write(0, field, fields[field][0])

    # 获取并写入数据段信息
    row = 1
    col = 0
    for row in range(1, len(results) + 1):
        for col in range(0, len(fields)):
            if  results[row-1][col] is None:
                sheet.write(row, col, u'%s' % '')
            else:
                sheet.write(row, col, u'%s' % results[row - 1][col])

    workbook.save(out_path)



host = "127.0.0.1"
port = 3306
user = 'root'
passwd = "he123123"
db = "test"
charset = "utf8"
print("mysql连接是否使用默认配置？y/n")
choose = input()
if operator.eq('n',choose) or operator.eq('n',choose):
    print("host:")
    host = input()
    print("port:")
    port = input()
    print("userName:")
    user = input()
    print("passWord:")
    passwd = input()
    print("databaseName:")
    db = input()
print("请输入Excel表格名称：")
sheet_name = input()
print("请粘贴您的SQL:")
sql = input()
print("指定盘符 C D E F:")
cd = input().upper()
out_path = r'F:/'+sheet_name+'.xls'
out_path = out_path.replace('F',cd)
print('输入路径为',out_path)
toExcel(host,port,user,passwd,db,charset,sql,sheet_name,out_path)
print("导出成功您的文件位置为：",out_path)