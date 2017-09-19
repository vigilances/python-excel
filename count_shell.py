import xlrd
import xlwt
input_month=input("Please input the month:")
wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)  
sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
sheet_count = wbk.add_sheet('sheet2', cell_overwrite_ok=True)
book = xlrd.open_workbook("file.xls")
#print("The number of worksheets is {0}".format(book.nsheets))
#print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
#print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
#print("Cell spilt_time7 is {0}".format(sh.cell_value(rowx=6, colx=3)))
#print(sh.col(0))
style=xlwt.XFStyle()
al = xlwt.Alignment()
al.horz = xlwt.Alignment.HORZ_CENTER 
al.vert = xlwt.Alignment.VERT_CENTER 
style.alignment = al
sheet.write(0,0,"姓名",style)
sheet.write(0,1,"日期",style)
sheet.write(0,2,"记录时间",style)
sheet.write(0,3,"加班时间",style)
sheet.write(0,4,"申请调休时间（小时）",style)
sheet.write(0,5,"累计（小时）",style)
sheet_count.write_merge(0,1,0,5,"技术一部"+input_month+"月加班餐补明细",style)
sheet_count.write(2,0,"姓名",style)
sheet_count.write(2,1,"日期",style)
sheet_count.write(2,2,"记录时间",style)
#sheet_count.write_merge(2,2,1,2,input_month+"月餐补天数",style)
#sheet_count.write(2,1,input_month+"月餐补天数",style)
sheet_count.write(2,3,"每次补助",style)
sheet_count.write(2,4,"合计补助",style)
sheet_count.write(2,5,"签字",style)
i=0
j=2
all_time=0
is_overtime={}
count_money=0

for rx in range(sh.nrows):  
    id_tag=sh.cell_value(rx,0)
    #第一个必须是打卡号最前面的那个，要不合并单元格会出问题
    #胡良恒
    if id_tag =="5":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        dates=[]
        print(name_info)   
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
               # sheet.write(i,0,name_info)
                sheet.write_merge(1,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+1,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+1,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+1,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
                #sheet.write(i,5,all_time)
                sheet.write_merge(1,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))      
        all_time=0
        count_money=0
        print("***********分割线************")
    #谢迪
    elif id_tag=="97":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)        
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
               # sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style) 
               # sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))
        all_time=0
        count_money=0
        print("***********分割线************")
    #吴仲春
    elif id_tag=="343":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info) 
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))
        all_time=0
        count_money=0
        print("***********分割线************")
    #屈喆
    elif id_tag=="371":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1           
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5  
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
               # sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,"屈喆",style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,"屈喆",style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
               # sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime["屈喆"]=str(len(dates))
                print("datetimes "+str(len(dates)))              
        all_time=0
        count_money=0
        print("***********分割线************")
    #胡飞
    elif id_tag=="406":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1   
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
               # sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
               # sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))             
        all_time=0
        count_money=0
        print("***********分割线************")
    #康子庄
    elif id_tag=="418":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1      
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))              
        all_time=0
        count_money=0
        print("***********分割线************")
    #张鑫
    elif id_tag=="22":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)    
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1      
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
                #sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
                #sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates=[]
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))              
        all_time=0
        count_money=0
        print("***********分割线************")
    #邓杰
    elif id_tag=="18":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)     
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1                
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))                
        all_time=0
        count_money=0
        print("***********分割线************")
    #刘波
    elif id_tag=="26":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)      
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))                            
        all_time=0
        count_money=0
        print("***********分割线************")
    #李俊葶
    elif id_tag=="35":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)  
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5 
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))               
        all_time=0
        count_money=0
        print("***********分割线************")
        #曹然
    elif id_tag=="82":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)    
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))             
        all_time=0
        count_money=0
        print("***********分割线************")
    #测试
    elif id_tag=="xx":
        #row_info=(sh.row(rx))
        name_info=sh.cell_value(rx,1)
        temp=i+1
        dates=[]
        print(name_info)    
        for cl in range(sh.ncols):
            #08:57\n18:26\n
            time_info=sh.cell_value(rx,cl)
            spilt_time=(time_info[-6:-4])
            if spilt_time=="":
                spilt_time+="0"
            #print(spilt_time)
            if(int(spilt_time)>=20):
                i+=1
                each_time=(int(time_info[-6:-4])-18)
                if(int(time_info[-3:-1])>=30):
                    each_time+=0.5
                all_time+=each_time
                count_money+=30
                date_info=str(sh.cell_value(2,cl))
                print(date_info)
              #  sheet.write(i,0,name_info)
                sheet.write_merge(temp,i,0,0,name_info,style) 
                sheet.write(i,1,(input_month+"."+date_info[-4:-2]),style)
                sheet.write(i,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write_merge(j+temp,i+2,0,0,name_info,style)
                sheet_count.write(i+2,1,(input_month+"."+date_info[-4:-2]),style,)
                sheet_count.write(i+2,2,"18:00-"+time_info[-6:-1],style)
                sheet_count.write(i+2,3,"30",style)
                sheet_count.write_merge(j+temp,i+2,4,4,count_money,style)
                sheet_count.write_merge(j+temp,i+2,5,5,"",style)
                sheet.write(i,3,each_time,style)
                sheet.write(i,4,each_time,style)
              #  sheet.write(i,5,all_time)
                sheet.write_merge(temp,i,5,5,all_time,style)
                dates.append(name_info)
                is_overtime[name_info]=str(len(dates))
                print("datetimes "+str(len(dates)))             
        all_time=0
        count_money=0
        print("***********分割线************")     
sheet_count.write(j+temp,0,"合计",style)
sheet_count.write(j+temp,1,i,style)
sheet_count.write(j+temp,2,"\\",style)
sheet_count.write(j+temp,3,"30",style)
sheet_count.write(j+temp,4,str(i*30),style)
sheet_count.write(j+temp,5,"\\",style)
wbk.save("D:\技术一部加班统计---"+input_month+"月.xls")
