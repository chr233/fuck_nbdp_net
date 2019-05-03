import requests
import re
import xlwt
import urllib3
from io import BytesIO
from bs4 import BeautifulSoup  

'''
By Chr_
'''

def download_exam():
    s=requests.session()
    workbook = xlwt.Workbook(encoding = 'utf-8')
    for j in range(34,100,1):
        print('当前进度%d/100' % j)
        txt=[]
        url='http://lib.nbdp.net/paper/%d.html'  % j
        html=s.get(url).content
        html=str(html,encoding='utf-8',errors='ignore')
        soup = BeautifulSoup(html,'lxml')  
        exams=soup.find_all(name='div',attrs={'s':'math3'})
        title=str(j)+soup.title.get_text()
        worksheet = workbook.add_sheet(title)
        exam=[(url,title)]
        for x in exams:
            try:
                #选择题
                tigan=x.find(name='p',attrs={'class':'pt1'}).get_text().strip()
                xuanxiang=[]
                for i in x.find_all(name='li'):
                    xuanxiang.append(i.get_text())
                try:
                    key=x.find(attrs={'class':'col-md-3 column xz'}).get_text()
                except AttributeError:
                    key='未找到答案'
                try:
                    img=x.find(name='img').attrs['src']
                    out={'tg':tigan,'xx':xuanxiang,'da':key,'tp':img}
                except AttributeError:
                    out={'tg':tigan,'xx':xuanxiang,'da':key}
                #print(out)
            except AttributeError:
                #填空题
                tigan=x.find(name='p').get_text().strip()
                try:
                    code=x.find(name='pre').get_text()
                    try:
                        img=x.find(name='img').attrs['src']
                        out={'tg':tigan,'dm':code,'tp':img}
                    except AttributeError:
                        out={'tg':tigan,'dm':code}
                except AttributeError:
                    jianda=x.find(name='span').get_text()
                    try:
                        img=x.find(name='img').attrs['src']
                        out={'tg':tigan,'jd':jianda,'tp':img}
                    except AttributeError:
                        out={'tg':tigan,'jd':jianda}
            finally:
                exam.append(out)
        sheetwriter(exam,worksheet)
        #workbook.save('i.xls')
    workbook.save('dump.xls')
    pass

def sheetwriter(list,sheetobj):
    sheetobj.write(0,1, xlwt.Formula('HYPERLINK("%s";"%s")' % list[0]))
    sheetobj.write(0,2, xlwt.Formula('HYPERLINK("https://blog.chrxw.com";"Generate By Chr_")'))
    sheetobj.col(0).width=0
    sheetobj.col(1).width=80*256
    sheetobj.col(2).width=40*256
    sheetobj.col(3).width=40*256
    sheetobj.col(4).width=40*256
    sheetobj.col(5).width=40*256
    _row=1
    for item in list:
        if 'tp' in item:#插图片
            url='http://lib.nbdp.net/'+item['tp']
            sheetobj.write(_row,6, xlwt.Formula('HYPERLINK("%s";"查看图片")' % url))
            pass
        if 'xx' in item:#选择题
            sheetobj.write(_row,1, label =item['tg'])
            sheetobj.write(_row,0, label =item['da'])
            try:
                sheetobj.write(_row,2, label =item['xx'][0])
                sheetobj.write(_row,3, label =item['xx'][1])
                sheetobj.write(_row,4, label =item['xx'][2])
                sheetobj.write(_row,5, label =item['xx'][3])
            except Exception:
                sheetobj.write(_row,7, label ='数据有误 ，请参照原网站')
            _row+=1
            continue
        if 'dm' in item:#填空题
            lines=item['tg'].splitlines(False)
            _row+=1
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            lines=item['dm'].splitlines(False)
            for line in lines:
                sheetobj.write(_row,1, label =line)
                _row+=1
            continue
        if 'jd' in item:#简答题
            sheetobj.write(_row,1, label =item['tg'])
            sheetobj.write(_row,0, label =item['jd'])
            _row+=1
    pass
        
if __name__=="__main__":
    download_exam()
