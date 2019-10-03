import xlsxwriter
import bs4
import requests


product = ['Animation_Designer', 'CAM_DATA_PREP', 'Die_Design','Die_Engineering','Electrode_Design','Engineering_Die_Wizard','Expression_Design_Logic','General_Packaging','Global_Shaping','KDA_Misc','Knowledge_Fusion','Measurement','Mechatronics','Mold_Wizard','Part_Family','Progressive_Die','Reuse','Ship_Design','Validation','Weld_Assistant']

baseline_previous = str(input('请输入previous baseline：'))
baseline_latest = str(input('请输入latest baseline：'))
name = str(input('请输入表单名字：'))

detail_url =  "http://cipgweb/autotest/kda/details.php?Build1=" + baseline_previous + "&submit_it=1&check_opt%5B0%5D=show_pr&check_opt%5B1%5D=show_new_fails&check_opt%5B2%5D=show_fails&check_opt%5B3%5D=show_not_run&check_opt%5B4%5D=show_new_passes&platform_opt%5B0%5D=win64&filter=NONE&type_filter=NONE&Build1=" + baseline_latest
compare_url = "http://cipgweb/autotest/kda/compare.php?Build1=" + baseline_previous + "&Build2=" + baseline_latest + "&submit_it=1&check_opt%5B%5D=show_pr&check_opt%5B%5D=show_new_fails&check_opt%5B%5D=show_fails&check_opt%5B%5D=show_not_run&check_opt%5B%5D=show_new_passes&platform_opt%5B%5D=win64&filter=NONE&type_filter=NONE"

#soup detail_url
res_detail = requests.get(detail_url)
bs_detail = bs4.BeautifulSoup(res_detail.text, 'html.parser') 

#get total
total = bs_detail.select('. td')
if len(total)>0:
    total_num = total[2].getText()

#get pass
passing = bs_detail.select('.pass td')
passing_num = passing[2].getText() #pass + %
    #passing_num = passing_all.split()[0]
#get pass rate
    #a = passing_all.split()[1]
    #passing_rate = a[1:-1]

#get fail
fail = bs_detail.select('.fail td')
if len(fail)>0:
    fail_all = fail[2].getText() #fail + %
    fail_num = fail_all.split()[0] #拆分str


#soup compare_url
res_compare = requests.get(compare_url)
bs_compare = bs4.BeautifulSoup(res_compare.text, 'html.parser')

#get Regression
regression = bs_compare.select('.regression td')
if len(regression)>0:
    regression_num = regression[2].getText()
    #regression_num = regression_all.split()[0]
else:
    regression_num = 0
    
#get New Add
newadd = bs_compare.select('.result td a')
if len(newadd)>0:
    newadd_num = newadd[0].getText()
else:
    newadd_num = 0
    
#get New Pass
newpass = bs_compare.select('.newpass td')
if len(newpass)>0:
    newpass_all = newpass[2].getText()
    newpass_num = newpass_all.split()[0]
else:
    newpass_num = 0
#get Not Run
notrun = bs_compare.select('.notrun td')
if len(notrun)>0:
    notrun_all = notrun[2].getText()
    notrun_num = notrun_all.split()[0]
else:
    notrun_num = 0

workbook = xlsxwriter.Workbook('D:\\report\\'+ name + '.xlsx')

worksheet = workbook.add_worksheet()

#set format:
wrap_format = workbook.add_format()
wrap = wrap_format.set_text_wrap()
bold_format = workbook.add_format()
bold = bold_format.set_bold()
#title_format = workbook.add_format({'bold':True})
#border = cellformat.set_border(1)

A1_text = 'Baseline:' + baseline_latest + '\nTotal:'+ str(total_num) + '\nOver All Pass Rate:' + str(passing_num) + '\nNew Added Test:' + str(newadd_num) + '\nNew Introduced failure:' + str(regression_num) +'\nTotal report: Click This'

worksheet.write('A1',A1_text,wrap)
worksheet.write('B1','Link',bold)
worksheet.write('C2','Product',bold)
worksheet.write('D2','Regression',bold)
worksheet.write('E2','New Pass',bold)
worksheet.write('F2','New Added',bold)
worksheet.write('G2','Pass',bold)
worksheet.write('H2','Fail',bold)
worksheet.write('I2','Not Runs',bold)
worksheet.write('J2','Total',bold)
worksheet.write('K2','Pass Rate',bold)
worksheet.write('L2','Compare to',bold)
worksheet.write('M2', baseline_previous)

worksheet.set_column(0,0,32)  #set A1 cell width
worksheet.set_row(0,105)      #set A1 cell hight
worksheet.set_column(2,2,23)  #set column C width
worksheet.set_column(11,11,27) #set column L width

#for row in range(2,11):
    

line = 2
row = 2

for i in product:
    detail_web_func = "http://cipgweb/autotest/kda/details.php?Build1=" + baseline_latest + "&Build2=" + baseline_latest + "&submit_it=1&check_opt%5B%5D=show_pr&check_opt%5B%5D=show_new_fails&check_opt%5B%5D=show_fails&check_opt%5B%5D=show_not_run&check_opt%5B%5D=show_new_passes&platform_opt%5B%5D=win64&filter=" + i + "&type_filter=NONE"
    res_func_detail = requests.get(detail_web_func)
    bs_func_detail = bs4.BeautifulSoup(res_func_detail.text, 'html.parser')
    compare_web_func = "http://cipgweb/autotest/kda/compare.php?Build1=" + baseline_previous + "&Build2=" + baseline_latest + "&submit_it=1&check_opt%5B%5D=show_pr&check_opt%5B%5D=show_new_fails&check_opt%5B%5D=show_fails&check_opt%5B%5D=show_not_run&check_opt%5B%5D=show_new_passes&platform_opt%5B%5D=win64&filter=" + i + "&type_filter=NONE"
    res_func_compare = requests.get(compare_web_func)
    bs_func_compare = bs4.BeautifulSoup(res_func_compare.text, 'html.parser')
    
    #get func total
    func_total = bs_func_detail.select('. td')
    if len(func_total)>0:
        func_total_num = func_total[2].getText()
    else:
        func_total_num = ''
        
    #get func pass
    func_passing = bs_func_detail.select('.pass td')
    if len(func_passing)>0:
        func_passing_all = func_passing[2].getText()#得到pass数值 + 百分比
        func_passing_num = func_passing_all.split()[0]#得到pass数值
    #get func pass rate
        a = func_passing_all.split()[1]
        func_passing_rate = a[1:-1] #得到pass百分比

    #get func fail
    func_fail = bs_func_detail.select('.fail td')
    if len(func_fail)>0:
        func_fail_all = func_fail[2].getText()
        func_fail_num = func_fail_all.split()[0]
    else:
        func_fail_num = ''
        
    #compare:
    #get func regression
    func_regression = bs_func_compare.select('.regression td')
    if len(func_regression)>0:
        func_regression_all = func_regression[2].getText()
        func_regression_num = func_regression_all.split()[0]
    else:
        func_regression_num = ''
        
    #get func new pass
    func_newpass = bs_func_compare.select('.newpass td')
    if len(func_newpass)>0:
        func_newpass_all = func_newpass[2].getText()
        func_newpass_num = func_newpass_all.split()[0]
    else:
        func_newpass_num = ''
        
    #get func notrun
    func_notrun = bs_func_compare.select('.notrun td')
    if len(func_notrun)>0:
        func_notrun_all = func_notrun[2].getText()
        func_notrun_num = func_notrun_all.split()[0]

    else:
        func_notrun_num = ''
        
    #get func new add 
    func_newadd = bs_func_compare.select('.result td a')
    if len(func_newadd)>0:
        func_newadd_num = func_newadd[0].getText()
    else:
        func_newadd_num = ''

    worksheet.write(line,row,i,bold)
    worksheet.write(line,row+1,func_regression_num,bold)
    worksheet.write(line,row+2,func_newpass_num,bold)
    worksheet.write(line,row+3,func_newadd_num,bold)
    worksheet.write(line,row+4,func_passing_num,bold)
    worksheet.write(line,row+5,func_fail_num,bold)
    worksheet.write(line,row+6,func_notrun_num,bold)
    worksheet.write(line,row+7,func_total_num,bold)
    worksheet.write(line,row+8,func_passing_rate,bold)
    if len(func_fail)>0:
        worksheet.write(line,row+9,'Detail')
    
    line=line+1


workbook.close()
