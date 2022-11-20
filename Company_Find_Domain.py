# @Author:jorge
import requests
import urllib3
import xlwt
urllib3.disable_warnings()

while True:
    Beian_Api = 'https://hlwicpfwc.miit.gov.cn/icpproject_query/api/icpAbbreviateInfo/queryByCondition'
    token = 'eyJ0eXBlIjoxLCJ1IjoiMDk4ZjZiY2Q0NjIxZDM3M2NhZGU0ZTgzMjYyN2I0ZjYiLCJzIjoxNjY4OTQ2ODU5NjQ2LCJlIjoxNjY4OTQ3MzM5NjQ2fQ.t4PXMBgPknSBLTjriJh99OI6BqvJqcuX-uXfoe7MR0E' #需要在浏览器中查看并粘贴到这里,token失效时间很短，要马上粘贴进去
    unitName = input("请输入你要查询的域名或公司名称：")

    header = {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/json',
        'Cookie': '__jsluid_s=437b0994d0cff2ac9bb8014e1ad20e3c',
        'Origin': 'https://beian.miit.gov.cn',
        'Referer': 'https://beian.miit.gov.cn/',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36',
        'token': token

    }
    data = {
        "pageNum":"",
        "pageSize":"",
        "unitName":unitName
    }

    res = requests.post(Beian_Api,headers=header,json=data,verify=False)
    response = res.json()

    if response['code'] == 200:
        page = response['params']['lastPage']   #页数
        total = response['params']['total']     #结果数
        print(f'查询完毕，共有{page}页，{total}条结果')
        print('开始存储结果到文件')
    else:
        print('token已过期，请刷新浏览器！！！')
        break
    data2 = {
        "pageNum": "",
        "pageSize": f"{total}",
        "unitName": unitName
    }

    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet(u'结果', cell_overwrite_ok=True)
    sheet.write(0, 0, '域名')
    sheet.write(0, 1, '备案名')
    sheet.write(0, 2, '更新时间')
    res2 = requests.post(Beian_Api,headers=header,json=data2,verify=False)
    response2 = res2.json()

    for i in range(total):
        domain = response2['params']['list'][i]['domain']  #域名
        serviceLicence = response2['params']['list'][i]['serviceLicence']  #备案名
        updateRecordTime = response2['params']['list'][i]['updateRecordTime']  #更新时间
        unitName = response2['params']['list'][i]['unitName']  #备案主体
        sheet.write(i+1,0,domain)
        sheet.write(i+1,1,serviceLicence)
        sheet.write(i+1,2,updateRecordTime)
        sheet.write(i+1,3,unitName)

    book.save(f'{unitName}域名备案结果.xls')
    break
print('结束')

