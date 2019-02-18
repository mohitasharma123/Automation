import xlsxwriter
import xlrd
from tkinter import *

book = xlrd.open_workbook('C:\\Users\\mohita\\Documents\\Test_Case_Automation\\columns_tlc.xlsx')
sheet = book.sheet_by_index(0)

master = Tk()
master.wm_title("Info Required")
  
def var_states():
   print("Receiving: %d,\nValidation: %d,\nICFlatten: %d,\nICLatest: %d,\nDPTeradata: %d,\nDPCam: %d,\nOperationalRaw: %d,\nDPRefinedDelta: %d,\nDPRefinedLatest: %d" % (var1.get(), var2.get(), var3.get(), var4.get(), var5.get(),var6.get(),var7.get(),var8.get(),var9.get()))
  
var1 = IntVar()
Checkbutton(master, text="Receiving", variable=var1).grid(row=2, column=1, sticky=W)
var2 = IntVar()
Checkbutton(master, text="Validation", variable=var2).grid(row=2, column=3, sticky=W)
var3 = IntVar()
Checkbutton(master, text= "IC Flatten", variable=var3).grid(row=2, column=5, sticky=W)
var4 = IntVar()
Checkbutton(master, text="IC Latest", variable=var4).grid(row=3, column=1, sticky=W)
var5 = IntVar()
Checkbutton(master, text="DP Teradata", variable=var5).grid(row=3, column=3, sticky=W)
var6 = IntVar()
Checkbutton(master, text="DP Cam", variable=var6).grid(row=3, column=5, sticky=W)
var7 = IntVar()
Checkbutton(master, text="Operational Raw", variable=var7).grid(row=4, column=1, sticky=W)
var8 = IntVar()
Checkbutton(master, text="DP Refined Delta", variable=var8).grid(row=4, column=3, sticky=W)
var9 = IntVar()
Checkbutton(master, text="DP Refined Latest", variable=var9).grid(row=4, column=5, sticky=W)
var10 = IntVar()
Checkbutton(master, text="Raw data Import", variable=var10).grid(row=1, column=1, sticky=W)
var11 = IntVar()
Checkbutton(master, text="Generic Data Import", variable=var11).grid(row=1, column=3, sticky=W)
var12 = IntVar()
Checkbutton(master, text="Delta Logic Applicable", variable=var12).grid(row=1, column=5, sticky=W)


Label(master, text="Datatopic").grid(row=0, column=1)
e1 = Entry(master)
e1.grid(row=0, column=2)
Label(master, text="HP ALM Path").grid(row=0, column=3)
e2 = Entry(master)
e2.grid(row=0, column=4)
Label(master, text="Output Excel Name").grid(row=0, column=5)
e3 = Entry(master)
e3.grid(row=0, column=6)

Button(master, text='OK', command=master.quit).grid(row=10, column=2, sticky=W, pady=4)
Button(master, text='Quit', command=quit).grid(row=10, column=4, sticky=W, pady=4)

mainloop()
receiving=var1.get()
validation=var2.get()
ICFlatten=var3.get()
ICLatest=var4.get()
DPTeradata=var5.get()
DPCam=var6.get()
OpRaw=var7.get()
DPRefDelta=var8.get()
DPRefLatest=var9.get()
RawDataImp=var10.get()
GenericImp=var11.get()
DtLogic=var12.get()
data_topic=e1.get()
path=e2.get()
file_name=e3.get()

workbook = xlsxwriter.Workbook('C:\\Users\\mohita\\Documents\\Test_Case_Automation\\%s.xlsx' % file_name)
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
worksheet.write_rich_string('A1', 'Layer', bold)
worksheet.write_rich_string('B1', 'Subject', bold)
worksheet.write_rich_string('C1', 'Test Case', bold)
worksheet.write_rich_string('D1', 'Type', bold)
worksheet.write_rich_string('E1', 'Step Name', bold)
worksheet.write_rich_string('F1', 'Descripition', bold)
worksheet.write_rich_string('G1', 'Expected result', bold)

merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter'})

border_format = workbook.add_format({
    'border': 1})


def rawDataFunc(i,j,k,l):    
    rawImp_user=sheet.cell(1,1).value
    freq=sheet.cell(1,3).value
    job_type=sheet.cell(1,4).value
    rawimp_hdfs=rawImp_user[:-2]

    if ( rawImp_user == "" or freq == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "Raw Import is selected but data in the column A is empty. Please quit and update columns.xlsx sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
    
    def rowFunc(i):
        rowA='A'
        rowB='B'
        rowC='C'
        rowD='D'
        rowE='E'
        rowG='G'
        rowA +=str(i)
        rowD +=str(i)
        rowB +=str(i)
        rowC +=str(i)
        rowE +=str(i)
        rowG +=str(i)
        return rowA,rowB,rowC,rowD,rowE,rowG

    rowX=i-1
    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(i)
    worksheet.write_rich_string(rowA,'kafka-storm')
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\kafka-storm')
    worksheet.write_rich_string(rowC,'TC-001_Validate when file is getting write on hdfs path then in each file record count must not increase 1500.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Connect to Devtest environment.') 
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(j)
    worksheet.write_rich_string(rowE,'2. Kinit-user.sh ',bold,'%s' % rawImp_user )
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(k)
    worksheet.write_rich_string(rowE,'3. Go to path /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(l)
    worksheet.write_rich_string(rowE,'4. Run wordcount command to check the count of file.')
    worksheet.write_rich_string(rowG,'when file is getting write on hdfs path then in each file record count must not increase 1500.')
    l=l+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\kafka-storm')
    worksheet.write_rich_string(rowC,'TC-002_Validate data file will be written on path : /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Connect to Devtest environment.')
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(j)
    worksheet.write_rich_string(rowE,'2. Kinit-user.sh ',bold,'%s' % rawImp_user )
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(k)
    worksheet.write_rich_string(rowE,'3. Go to path /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(l)
    worksheet.write_rich_string(rowE,'4. Check file is getting written at this path.')
    worksheet.write_rich_string(rowG,'data file will be written on path: /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    rowY=l-1
    worksheet.merge_range(rowX, 0, rowY, 0, 'kafka-storm', merge_format)
    l=l+4

    rowX=i-1
    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(i)
    worksheet.write_rich_string(rowA,'RawDataImport')
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\RawDataImport')
    worksheet.write_rich_string(rowC,'TC-001_Validate rawdataimport job for ',bold,'%s' % data_topic,' execute every ',bold,'%s' % freq,' hour/s.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % rawImp_user)
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for',bold,' %s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(l)
    worksheet.write_rich_string(rowE,'4. Check workflow executing every ',bold,'%s' % freq,' hour/s')
    worksheet.write_rich_string(rowG,'rawdataimport job for OCIS Contracts execute every ',bold,'%s' % freq,' hour/s.')
    l=l+4


    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\RawDataImport')
    worksheet.write_rich_string(rowC,'TC-002_Validate only 30 files will be moved from storm inbox path to rawpipeline path.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Connect to Devtest environment.')
    i=i+4
    

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(j)
    worksheet.write_rich_string(rowE,'2. Kinit-user.sh ',bold,'%s' % rawImp_user )
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(k)
    worksheet.write_rich_string(rowE,'3. Go to path /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(l)
    worksheet.write_rich_string(rowE,'4. Check only 30 files should be moved.')
    worksheet.write_rich_string(rowG,'only 30 files will be moved from storm inbox path to rawpipeline path.')
    l=l+4


    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\RawDataImport')
    worksheet.write_rich_string(rowC,'TC-003_Validate file will be moved to at rawpipeline path in directory structure path. Day and time value will be execution day and time of rawdataimport job.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Connect to Devtest environment.')
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(j)
    worksheet.write_rich_string(rowE,'2. Kinit-user.sh ',bold,'%s' % rawImp_user )
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(k)
    worksheet.write_rich_string(rowE,'3. Go to path /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/inbox.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc(l)
    worksheet.write_rich_string(rowE,'4. Check after successful execution, all files moved in same processing day and time')
    worksheet.write_rich_string(rowG,'file will be moved to at rawpipeline path in directory structure like /datasets/encrypted/',bold,'%s' % rawimp_hdfs,'/rawpipeline/day=${YEAR}-${MONTH}-${DAY}/time=${HOUR}-${MINUTE}.Day and time value will be execution day and time of rawdataimport job.')
    rowY=l-1
    a=l
    l=l+4
    m=l+1
    worksheet.merge_range(rowX, 0, rowY, 0, 'RawDataImport', merge_format)
##    worksheet.conditional_format( rowX, 0, rowY, 6 , { 'format' : border_format} )
##    worksheet.set_column(rowX, 0, rowY, 6, border_format)
    return i,j,k,l,m,a

def genDataFunc(i,j,k,l):
    genImp_user=sheet.cell(2,1).value
    file_name=sheet.cell(2,5).value
    job_type=sheet.cell(2,4).value
    file_typ=file_name[-4:]
    freq=sheet.cell(2,3).value
    genimp_hdfs=genImp_user[:-2]

    if ( genImp_user == "" or file_name =="" or freq =="" or job_type =="" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "Generic Import is selected but data for the generic import in input excel is empty. Please quit and update input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()

   
    def rowFunc_1(i):
        rowA='A'
        rowB='B'
        rowC='C'
        rowD='D'
        rowE='E'
        rowG='G'
        rowA +=str(i)
        rowD +=str(i)
        rowB +=str(i)
        rowC +=str(i)
        rowE +=str(i)
        rowG +=str(i)
        return rowA,rowB,rowC,rowD,rowE,rowG

    rowX=i-1
    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)        
    worksheet.write_rich_string(rowA,'Generic data import')
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-001_Validate GenericDataImport job for ',bold,'%s' % data_topic,' execute every ',bold,'%s' % freq,' hour/s')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check workflow executing every ',bold,'%s' % freq,' hour/s')
    worksheet.write_rich_string(rowG,'GenericDataImport job for ',bold,'%s' % data_topic,' execute every ',bold,'%s' % freq,' hour/s')
    l=l+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-002_Validate ',bold,'%s' % file_name,' and ',bold,'%s.metadata' % file_name,' file present in /datasets/encrypted/',bold,'%s' % genimp_hdfs,'/landingzone path')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+2

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Validate ',bold,'%s' % file_typ,' and .metadata files are present in landing zone')
    worksheet.write_rich_string(rowG,'Validate ',bold,'%s' % file_name,' and ',bold,'%s.metadata' % file_name,' file present in /datasets/encrypted/',bold,'%s' % genimp_hdfs,'/landingzone path')    
    j=j+2
    k=k+2
    l=l+2

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-003_Validate total number of records in ',bold,'%s' % file_typ,' file and record count in metadata file matches.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Validate ',bold,'%s' % file_typ,' and .metadata files are present in landing zone')
    worksheet.write_rich_string(rowG,'Validate ',bold,'%s' % file_name,' and ',bold,'%s.metadata' % file_name,' file present in /datasets/encrypted/',bold,'%s' % genimp_hdfs,'/landingzone path')
    j=j+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Validate the count in ',bold,'%s' % file_typ,' and .metadata files')
    worksheet.write_rich_string(rowG,'Validate total number of records in *',bold,'%s' % file_typ,' file and record count in metadata file matches.')
    k=k+3
    l=l+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-004_Validate if total number of records in *',bold,'%s' % file_typ,' file and record count in metadata file doesn’t matches then workflow gets failed.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Validate ',bold,'%s' % file_typ,' and .metadata files are present in landing zone')
    worksheet.write_rich_string(rowG,'Validate ',bold,'%s' % file_name,' and ',bold,'%s.metadata' % file_name,' file present in /datasets/encrypted/',bold,'%s' % genimp_hdfs,'/landingzone path')
    j=j+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    worksheet.write_rich_string(rowG,'Validate if total number of records in *',bold,'%s' % file_typ,' file and record count in metadata file doesn’t matches then workflow gets failed.')
    k=k+3
    l=l+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-005_Validate that the import job fails if there are no files in landingzone path')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Validate ',bold,'%s' % file_typ,' and .metadata files are not present in landing zone')
    worksheet.write_rich_string(rowG,'Validate ',bold,'%s' % file_name,' and ',bold,'%s.metadata' % file_name,' files are not present in /datasets/encrypted/',bold,'%s' % genimp_hdfs,'/landingzone path')
    j=j+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    worksheet.write_rich_string(rowG,'Validate that the import job fails as there are no files in landingzone path')
    k=k+3
    l=l+3

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % path,'\GenericDataImport')
    worksheet.write_rich_string(rowC,'TC-006_Validate file will be moved to at rawpipeline path in directory structure path. Day and time value will be execution day and time of GenericDataImport job.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % genImp_user)
    i=i+4 

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check after successful execution, all files moved in same processing day and time')
    worksheet.write_rich_string(rowG,'file will be moved to at rawpipeline path in directory structure like /datasets/',bold,'%s' % genimp_hdfs,'/rawpipeline/day=${YEAR}-${MONTH}-${DAY}/time=${HOUR}-${MINUTE}.Day and time value will be execution day and time of GenericDataImport job.')
    a=l
    rowY=l-1
    l=l+4
    m=l+1
    worksheet.merge_range(rowX, 0, rowY, 0, 'Generic data import', merge_format)
    return i,j,k,l,m,a


def genericCaseFunc(i,j,k,l,m,a,user,table,freq,layer,layer_path,db,p,job_type) :

      def rowFunc_1(i):
           rowA='A'
           rowB='B'
           rowC='C'
           rowD='D'
           rowE='E'
           rowG='G'
           rowA +=str(i)
           rowD +=str(i)
           rowB +=str(i)
           rowC +=str(i)
           rowE +=str(i)
           rowG +=str(i)
           return rowA,rowB,rowC,rowD,rowE,rowG

      table_final=table[-4:]
      table_final=table_final.upper()
      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
      q=format(p, '03')
      worksheet.write_rich_string(rowA,bold,'%s' % layer,' Layer')
      worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
      worksheet.write_rich_string(rowC,'TC-%s_Validate the workflow is successful and after successful execution data gets populated in ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' table.')
      worksheet.write_rich_string(rowD,'Manual')
      worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
      p=p+1
      i=i+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
      worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
      j=j+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
      worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
      worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
      k=k+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
      worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
      l=l+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
      worksheet.write_rich_string(rowE,'5. Check the workflow is successful and data inserted in processing day and time in ',bold,'%s' % table,' table.')
      worksheet.write_rich_string(rowG,'Workflow should be successful and after successful execution data gets populated in ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table.')
      m=m+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
      q=format(p, '03')
      worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
      worksheet.write_rich_string(rowC,'TC-%s_Validate the record count in ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' table matches with the source layer/table as per logic')
      worksheet.write_rich_string(rowD,'Manual')
      worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
      p=p+1
      i=i+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
      worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
      j=j+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
      worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
      worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
      k=k+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
      worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
      l=l+5
      
      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
      worksheet.write_rich_string(rowE,'5. Check the record count in ',bold,'%s' % table,' table.')
      worksheet.write_rich_string(rowG,'After successful execution,the record count in ',bold,'%s' % table,' table should match with the source layer/table as per logic')
      m=m+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
      q=format(p, '03')
      worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
      worksheet.write_rich_string(rowC,'TC-%s_Validate the data populated in the columns of ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' table is correctly mapped and populated.')
      worksheet.write_rich_string(rowD,'Manual')
      worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
      p=p+1
      i=i+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
      worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
      j=j+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
      worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
      worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
      k=k+5

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
      worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
      l=l+5
      
      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
      worksheet.write_rich_string(rowE,'5. Check the data populated in the columns of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table is correctly mapped and populated.')
      worksheet.write_rich_string(rowG,'The data populated in the columns of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table should be correctly mapped and populated.')
      m=m+5

      if (table_final != '_STG'):
         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
         q=format(p, '03')
         worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
         worksheet.write_rich_string(rowC,'TC-%s_Validate the data populated in the Watermark fields of ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' table is correctly calculated and populated.')
         worksheet.write_rich_string(rowD,'Manual')
         worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
         p=p+1
         i=i+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
         worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
         j=j+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
         worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
         worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
         k=k+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
         worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
         l=l+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
         worksheet.write_rich_string(rowE,'5. Check the data populated in the watermark fields of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table is correctly calculated and populated.')
         worksheet.write_rich_string(rowG,'The data populated in the watermark fields of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table should be correctly calculated and populated.')
         m=m+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
         q=format(p, '03')
         worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
         worksheet.write_rich_string(rowC,'TC-%s_Validate in hash_cde column of ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' table, hash_cde value is correctly calculated and populated.')
         worksheet.write_rich_string(rowD,'Manual')
         worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
         p=p+1
         i=i+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
         worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
         j=j+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
         worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
         worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
         k=k+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
         worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
         l=l+5

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
         worksheet.write_rich_string(rowE,'5. Check in hash_cde column of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table, hash_cde value is correctly calculated and populated.')
         worksheet.write_rich_string(rowG,'The data populated in the hash_cde column of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' table should be correctly calculated and populated.')
         m=m+5

         
      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
      q=format(p, '03')
      worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
      worksheet.write_rich_string(rowC,'TC-%s_Validate the DDL and column data types of ' % q,bold,'%s' % layer,' layer ',bold,'%s' % table,' are as per the wiki design.')
      worksheet.write_rich_string(rowD,'Manual')
      worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
      p=p+1
      i=i+3

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
      worksheet.write_rich_string(rowE,'2. connect to hive console and to ',bold,'%s' % db,' db.')
      j=j+3

      rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
      worksheet.write_rich_string(rowE,'3. Check the DDL and column data types of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' are as per the wiki design.')
      worksheet.write_rich_string(rowG,'The DDL and column data types of ',bold,'%s' % layer,' layer ',bold,'%s' % table,' should be as per the wiki design.')
      k=k+3
      l=l+3      
      m=m+3

      if( layer != 'DP Teradata' and layer != 'DP CAM' ):
         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
         q=format(p, '03')
         worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
         worksheet.write_rich_string(rowC,'TC-%s_Validate that the workflow is triggered at the defined intervals based on frequency i.e ' % q,bold,'%s' %freq,' hour/s in a long run')
         worksheet.write_rich_string(rowD,'Manual')
         worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
         p=p+1
         i=i+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
         worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
         j=j+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
         worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
         worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
         k=k+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
         worksheet.write_rich_string(rowE,'4. Check that the workflow is triggered at the defined intervals based on frequency i.e ',bold,'%s' %freq,' hour/s in a long run')
         worksheet.write_rich_string(rowG,'The workflow should be triggered at the defined intervals based on frequency i.e ',bold,'%s' %freq,' hour/s in a long run')
         l=l+4
         m=l+1

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
         worksheet.write_rich_string(rowB,bold,'%s' % layer_path)
         q=format(p, '03')
         worksheet.write_rich_string(rowC,'TC-%s_Validate the runtime of workflow is in the expected range while performing performance testing of ' % q,bold,'%s' % layer,' layer.')
         worksheet.write_rich_string(rowD,'Manual')
         worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % user )
         p=p+1
         i=i+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
         worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
         j=j+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
         worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
         worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
         k=k+4

         rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
         worksheet.write_rich_string(rowE,'4. Check the runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
         worksheet.write_rich_string(rowG,'The runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
         a=l
         l=l+4
         m=l+1

      return i,j,k,l,m,a,p

 
def receivingFunc(i,j,k,l,m,a):

    rec_user=sheet.cell(3,1).value
    rec_table=sheet.cell(3,2).value
    rec_freq=sheet.cell(3,3).value
    job_type=sheet.cell(3,4).value
    layer="Receiving"
    rec_path=path
    rec_path += '\Receiving Layer'
    db="bdpdt_receiving"
    p=1

    if ( rec_user == "" or rec_table == "" or rec_freq == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "Receiving Layer is selected but data for the receiving layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()

    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,rec_user,rec_table,rec_freq,layer,rec_path,db,p,job_type)

    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0, 'Receiving Layer', merge_format)
    
    return i,j,k,l,m,a

def validationFunc(i,j,k,l,m,a):

    val_user=sheet.cell(4,1).value
    val_table=sheet.cell(4,2).value
    val_freq=sheet.cell(4,3).value
    job_type=sheet.cell(4,4).value
    layer="Validation"
    val_path=path
    val_path += '\Validation Layer'
    db="bdpdt_validation"
    p=1

    if ( val_user == "" or val_table == "" or val_freq == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "Validation Layer is selected but data for the validation layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()

    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,val_user,val_table,val_freq,layer,val_path,db,p,job_type)

    def rowFunc_1(i):
       rowA='A'
       rowB='B'
       rowC='C'
       rowD='D'
       rowE='E'
       rowG='G'
       rowA +=str(i)
       rowD +=str(i)
       rowB +=str(i)
       rowC +=str(i)
       rowE +=str(i)
       rowG +=str(i)
       return rowA,rowB,rowC,rowD,rowE,rowG
      

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    q=format(p, '03')
    worksheet.write_rich_string(rowB,bold,'%s' % val_path)
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the dedup logic of 7 days is applied and duplicate records within the 7 days are discarded and not inserted in validation ' % q,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % val_user )
    p=p+1
    i=i+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
    l=l+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
    worksheet.write_rich_string(rowE,'5. Check that the dedup logic of 7 days is applied and duplicate records within the 7 days are discarded and not inserted in validation ' ,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowG,'The dedup logic of 7 days should be applied and duplicate records within the 7 days should be discarded and not inserted in validation ' ,bold,'%s' % val_table,' table')
    m=m+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    q=format(p, '03')
    worksheet.write_rich_string(rowB,bold,'%s' % val_path)
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the dedup logic of 7 days is applied and duplicate records after the 7 days are not discarded and inserted in validation ' % q,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % val_user )
    p=p+1
    i=i+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
    l=l+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
    worksheet.write_rich_string(rowE,'5. Check that the dedup logic of 7 days is applied and duplicate records after the 7 days are discarded and inserted in validation ' ,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowG,'The dedup logic of 7 days should be applied and duplicate records after the 7 days should be discarded and inserted in validation ' ,bold,'%s' % val_table,' table')
    m=m+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    q=format(p, '03')
    worksheet.write_rich_string(rowB,bold,'%s' % val_path)
    worksheet.write_rich_string(rowC,'TC-%s_Validate that if any of the column of natural key is null or blank then the record is discarded and not inserted in validation ' % q,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % val_user )
    p=p+1
    i=i+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
    l=l+5

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
    worksheet.write_rich_string(rowE,'5. Check that if any of the column of natural key is null or blank then the record should be discarded and not inserted in validation ' ,bold,'%s' % val_table,' table')
    worksheet.write_rich_string(rowG,'If any of the column of natural key is null or blank then the record should be discarded and not inserted in validation ' ,bold,'%s' % val_table,' table')
    m=m+5
    
    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0, 'Validation Layer', merge_format)

    return i,j,k,l,m,a

def icFlattenFunc(i,j,k,l,m,a):
      ICFlatten_user=sheet.cell(5,1).value
      ICFlatten_table=sheet.cell(5,2).value
      ICFlatten_freq=sheet.cell(5,3).value
      job_type=sheet.cell(5,4).value
      layer="IC Flatten"
      ICFlatten_path=path
      ICFlatten_path += '\IC Flatten Layer'
      db="bdpdt_ic"
      p=1

      if ( ICFlatten_freq == 0 or ICFlatten_user == "" or ICFlatten_table == "" or job_type == "" ):
         pop=Tk()
         T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
         T.insert(END, "IC Flatten Layer is selected but data for the ic flatten layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
         T.config(state=DISABLED) # forbid text edition
         button = Button(pop, text="Quit", command=quit)
         T.window_create(INSERT, window=button)
         T.pack()
         mainloop()

      rowX=i-1
      i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,ICFlatten_user,ICFlatten_table,ICFlatten_freq,layer,ICFlatten_path,db,p,job_type)

      rowY=m-6
      worksheet.merge_range(rowX, 0, rowY, 0, 'IC Flatten Layer', merge_format)

      return i,j,k,l,m,a

def icLatestFunc(i,j,k,l,m,a):

    ICLatest_user=sheet.cell(6,1).value
    ICLatest_table=sheet.cell(6,2).value
    ICLatest_freq=sheet.cell(6,3).value
    job_type=sheet.cell(6,4).value
    layer="IC Latest"
    ICLatest_path=path
    ICLatest_path += '\IC Latest Layer'
    db="bdpdt_ic"
    p=1

    if ( ICLatest_user == "" or ICLatest_table == "" or ICLatest_freq == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "IC Latest Layer is selected but data for the ic latest layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,ICLatest_user,ICLatest_table,ICLatest_freq,layer,ICLatest_path,db,p,job_type)

    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0, 'IC Latest Layer', merge_format)
    
    return i,j,k,l,m,a

def dpTeradataFunc(i,j,k,l,m,a,DtLogic):
    DPTeradata_user=sheet.cell(7,1).value
    DPTeradata_table=sheet.cell(7,2).value
    DPTeradata_freq=sheet.cell(7,3).value
    job_type=sheet.cell(7,4).value
    layer="DP Teradata"
    DPTeradata_path=path
    DPTeradata_path += '\DP Teradata Layer'
    db="bdpdt_delivery_prep"
    p=1

    if ( DPTeradata_freq == "" or DPTeradata_user == "" or DPTeradata_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP Teradata Layer is selected but data for the DP Teradata layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPTeradata_user,DPTeradata_table,DPTeradata_freq,layer,DPTeradata_path,db,p,job_type)

    def rowFunc_1(i):
        rowA='A'
        rowB='B'
        rowC='C'
        rowD='D'
        rowE='E'
        rowG='G'
        rowA +=str(i)
        rowD +=str(i)
        rowB +=str(i)
        rowC +=str(i)
        rowE +=str(i)
        rowG +=str(i)
        return rowA,rowB,rowC,rowD,rowE,rowG

    if (DtLogic == 1):
           
       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPTeradata_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when new records are received then status should be CREATE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradata_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when new records are received then status should be CREATE')  
       worksheet.write_rich_string(rowG,'The status of records should be CREATE when new records are received')
       m=m+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPTeradata_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when update on the existing record then status should be UPDATE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradata_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when update on the existing record then status should be UPDATE')  
       worksheet.write_rich_string(rowG,'The status of records should be UPDATE when there is update in existing record')
       m=m+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPTeradata_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when no update on the existing record (records already present in the previous partition)then status should be DELETE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradata_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when no update on the existing record (records already present in the previous partition)then status should be DELETE')  
       worksheet.write_rich_string(rowG,'The status of records should be DELETE when there is no update on the existing record (records already present in the previous partition)')
       m=m+5


    DPTeradataSTG_user=sheet.cell(8,1).value
    DPTeradataSTG_table=sheet.cell(8,2).value
    DPTeradataSTG_freq=sheet.cell(8,3).value
    job_type=sheet.cell(8,4).value
    layer="DP Teradata"
    DPTeradataSTG_path=path
    DPTeradataSTG_path += '\DP Teradata Layer'
    db="bdpdt_delivery_prep"
    
    if ( DPTeradataSTG_freq == "" or DPTeradataSTG_user == "" or DPTeradataSTG_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP Teradata Stage Layer is selected but data for the DP Teradata Stage layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()

    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPTeradataSTG_user,DPTeradataSTG_table,DPTeradataSTG_freq,layer,DPTeradataSTG_path,db,p,job_type)

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    q=format(p, '03')
    worksheet.write_rich_string(rowB,bold,'%s' % DPTeradataSTG_path)
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the workflow is triggered at the defined intervals based on frequency i.e ' % q,bold,'%s' %DPTeradataSTG_freq,' hour/s in a long run')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradataSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the workflow is triggered at the defined intervals based on frequency i.e ',bold,'%s' %DPTeradataSTG_freq,' hour/s in a long run')
    worksheet.write_rich_string(rowG,'The workflow should be triggered at the defined intervals based on frequency i.e ',bold,'%s' %DPTeradataSTG_freq,' hour/s in a long run')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPTeradataSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate the runtime of workflow is in the expected range while performing performance testing of ' % q,bold,'%s' % layer,' layer.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradataSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check the runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
    worksheet.write_rich_string(rowG,'The runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPTeradataSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the sqoop workflow is successful and data is sqooped to EDW Stage table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradataSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the sqoop workflow is successful and data is sqooped to EDW Stage table')
    worksheet.write_rich_string(rowG,'The sqoop workflow should be successful and data should be sqooped to EDW Stage table')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPTeradataSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the record count of data sqooped to EDW stage table matches with USP Staging table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradataSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the record count of data sqooped to EDW stage table matches with USP Staging table')
    worksheet.write_rich_string(rowG,'The record count of data sqooped to EDW stage table should match with USP Staging table')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPTeradataSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the data is populated correctly for each field in EDW Stage table from USP Staging table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPTeradataSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the data is populated correctly for each field in EDW Stage table from USP Staging table')
    worksheet.write_rich_string(rowG,'The data should be populated correctly for each field in EDW Stage table from USP Staging table')
    a=l
    l=l+4
    m=l+1
    
    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0, 'DP Teradata Layer', merge_format)
       
    return i,j,k,l,m,a

def dpCamFunc(i,j,k,l,m,a,DtLogic):
    DPCam_user=sheet.cell(9,1).value
    DPCam_table=sheet.cell(9,2).value
    DPCam_freq=sheet.cell(9,3).value
    job_type=sheet.cell(9,4).value
    layer="DP CAM"
    DPCam_path=path
    DPCam_path += '\DP CAM Layer'
    db="bdpdt_delivery_prep"
    p=1

    if ( DPCam_freq == "" or DPCam_user == "" or DPCam_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP CAM Layer is selected but data for the DP CAM layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPCam_user,DPCam_table,DPCam_freq,layer,DPCam_path,db,p,job_type)

    def rowFunc_1(i):
        rowA='A'
        rowB='B'
        rowC='C'
        rowD='D'
        rowE='E'
        rowG='G'
        rowA +=str(i)
        rowD +=str(i)
        rowB +=str(i)
        rowC +=str(i)
        rowE +=str(i)
        rowG +=str(i)
        return rowA,rowB,rowC,rowD,rowE,rowG

    if (DtLogic == 1):
       
       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPCam_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when new records are received then status should be CREATE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCam_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when new records are received then status should be CREATE')  
       worksheet.write_rich_string(rowG,'The status of records should be CREATE when new records are received')
       m=m+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPCam_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when update on the existing record then status should be UPDATE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCam_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when update on the existing record then status should be UPDATE')  
       worksheet.write_rich_string(rowG,'The status of records should be UPDATE when there is update in existing record')
       m=m+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
       q=format(p, '03')
       worksheet.write_rich_string(rowB,bold,'%s' % DPCam_path)
       worksheet.write_rich_string(rowC,'TC-%s_Validate that when no update on the existing record (records already present in the previous partition)then status should be DELETE' % q)
       worksheet.write_rich_string(rowD,'Manual')
       worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCam_user )
       p=p+1
       i=i+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
       worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
       j=j+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
       worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
       worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
       k=k+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
       worksheet.write_rich_string(rowE,'4. connect to hive console and to ',bold,'%s' % db,' db.')
       l=l+5

       rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(m)
       worksheet.write_rich_string(rowE,'5. Validate that when no update on the existing record (records already present in the previous partition)then status should be DELETE')  
       worksheet.write_rich_string(rowG,'The status of records should be DELETE when there is no update on the existing record (records already present in the previous partition)')
       m=m+5

    DPCamSTG_user=sheet.cell(10,1).value
    DPCamSTG_table=sheet.cell(10,2).value
    DPCamSTG_freq=sheet.cell(10,3).value
    job_type=sheet.cell(10,4).value
    layer="DP CAM"
    DPCamSTG_path=path
    DPCamSTG_path += '\DP CAM Layer'
    db="bdpdt_delivery_prep"

    if ( DPCamSTG_freq == "" or DPCamSTG_user == "" or DPCamSTG_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP CAM Stage Layer is selected but data for the DP CAM Stage layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()

    
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPCamSTG_user,DPCamSTG_table,DPCamSTG_freq,layer,DPCamSTG_path,db,p,job_type)

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    q=format(p, '03')
    worksheet.write_rich_string(rowB,bold,'%s' % DPCamSTG_path)
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the workflow is triggered at the defined intervals based on frequency i.e ' % q,bold,'%s' %DPCamSTG_freq,' hour/s in a long run')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCamSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the workflow is triggered at the defined intervals based on frequency i.e ',bold,'%s' %DPCamSTG_freq,' hour/s in a long run')
    worksheet.write_rich_string(rowG,'The workflow should be triggered at the defined intervals based on frequency i.e ',bold,'%s' %DPCamSTG_freq,' hour/s in a long run')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPCamSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate the runtime of workflow is in the expected range while performing performance testing of ' % q,bold,'%s' % layer,' layer.')
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCamSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check the runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
    worksheet.write_rich_string(rowG,'The runtime of workflow is in the expected range while performing performance testing of ',bold,'%s' % layer,' layer.')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPCamSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the sqoop workflow is successful and data is sqooped to CAM Stage table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCamSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the sqoop workflow is successful and data is sqooped to CAM Stage table')
    worksheet.write_rich_string(rowG,'The sqoop workflow should be successful and data should be sqooped to CAM Stage table')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPCamSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the record count of data sqooped to CAM stage table matches with USP Staging table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCamSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the record count of data sqooped to CAM stage table matches with USP Staging table')
    worksheet.write_rich_string(rowG,'The record count of data sqooped to CAM stage table should match with USP Staging table')
    l=l+4
    m=l+1

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(i)
    worksheet.write_rich_string(rowB,bold,'%s' % DPCamSTG_path)
    q=format(p, '03')
    worksheet.write_rich_string(rowC,'TC-%s_Validate that the data is populated correctly for each field in CAM Stage table from USP Staging table' % q)
    worksheet.write_rich_string(rowD,'Manual')
    worksheet.write_rich_string(rowE,'1. Kinit-user.sh ',bold,'%s' % DPCamSTG_user )
    p=p+1
    i=i+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(j)
    worksheet.write_rich_string(rowE,'2. Submit ', bold,'%s' % job_type,' job for ',bold,'%s' % data_topic)
    j=j+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(k)
    worksheet.write_rich_string(rowE,'3. Monitor workflow in oozie console.')
    worksheet.write_rich_string(rowG,'Oozie workflow should gets executed successfully.')
    k=k+4

    rowA,rowB,rowC,rowD,rowE,rowG=rowFunc_1(l)
    worksheet.write_rich_string(rowE,'4. Check that the data is populated correctly for each field in CAM Stage table from USP Staging table')
    worksheet.write_rich_string(rowG,'The data should be populated correctly for each field in CAM Stage table from USP Staging table')
    a=l
    l=l+4
    m=l+1
    
    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0,"DP CAM Layer", merge_format)
       
    return i,j,k,l,m,a

def opRawFunc(i,j,k,l,m,a):
    OperationalRaw_user=sheet.cell(11,1).value
    OperationalRaw_table=sheet.cell(11,2).value
    OperationalRaw_freq=sheet.cell(11,3).value
    job_type=sheet.cell(11,4).value
    layer="Operational Raw"
    OperationalRaw_path=path
    OperationalRaw_path += '\Operational Raw Layer'
    db="bdpdt_raw"
    p=1

    if ( OperationalRaw_freq == "" or OperationalRaw_user == "" or OperationalRaw_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "Operational Raw Layer is selected but data for the Operational raw layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,OperationalRaw_user,OperationalRaw_table,OperationalRaw_freq,layer,OperationalRaw_path,db,p,job_type)

    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0,"Operational Raw Layer", merge_format)
    
    return i,j,k,l,m,a

def dpRefDelFunc(i,j,k,l,m,a):
    DPRefinedDelta_user=sheet.cell(12,1).value
    DPRefinedDelta_table=sheet.cell(12,2).value
    DPRefinedDelta_freq=sheet.cell(12,3).value
    job_type=sheet.cell(12,4).value
    layer="DP Refined Delta"
    DPRefinedDelta_path=path
    DPRefinedDelta_path += '\DP Refined Delta Layer'
    db="bdpdt_refined"
    p=1

    if ( DPRefinedDelta_freq == "" or DPRefinedDelta_user == "" or DPRefinedDelta_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP Refined Delta Layer is selected but data for the DP Refined Delta layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPRefinedDelta_user,DPRefinedDelta_table,DPRefinedDelta_freq,layer,DPRefinedDelta_path,db,p,job_type)

    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0,"DP Refined Delta Layer", merge_format)
    
    return i,j,k,l,m,a

def dpRefLtstFunc(i,j,k,l,m,a):
    DPRefinedLatest_user=sheet.cell(13,1).value
    DPRefinedLatest_table=sheet.cell(13,2).value
    DPRefinedLatest_freq=sheet.cell(13,3).value
    job_type=sheet.cell(13,4).value
    layer="DP Refined Latest"
    DPRefinedLatest_path=path
    DPRefinedLatest_path += '\DP Refined Latest Layer'
    db="bdpdt_refined"
    p=1

    if ( DPRefinedLatest_freq == "" or DPRefinedLatest_user == "" or DPRefinedLatest_table == "" or job_type == "" ):
       pop=Tk()
       T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
       T.insert(END, "DP Refined Latest Layer is selected but data for the DP Refined Latest layer in input excel is empty. Please quit and update the input excel sheet and rerun\n")
       T.config(state=DISABLED) # forbid text edition
       button = Button(pop, text="Quit", command=quit)
       T.window_create(INSERT, window=button)
       T.pack()
       mainloop()
       
    rowX=i-1
    i,j,k,l,m,a,p=genericCaseFunc(i,j,k,l,m,a,DPRefinedLatest_user,DPRefinedLatest_table,DPRefinedLatest_freq,layer,DPRefinedLatest_path,db,p,job_type)

    rowY=m-6
    worksheet.merge_range(rowX, 0, rowY, 0,"DP Refined Latest Layer", merge_format)
    
    return i,j,k,l,m,a

if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 0 and DPCam == 0 and OpRaw == 0 and DPRefDelta == 0 and DPRefLatest == 0 or data_topic == "" or path == ""):
   pop=Tk()
   T = Text(pop, height=5, width=50, bg='lightgrey', relief='flat')
   T.insert(END, "Either Checkbox is not selected or Data topic or HP ALM Path is not provided. Please rerun the code and provide the appropriate inputs\n")
   T.config(state=DISABLED) # forbid text edition
   button = Button(pop, text="Quit", command=quit)
   T.window_create(INSERT, window=button)
   T.pack()
   mainloop()

if (RawDataImp == 1 ):
    i=2
    j=3
    k=4
    l=5
    i,j,k,l,m,a=rawDataFunc(i,j,k,l)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 1 ):
    T=1
    i=2
    j=3
    k=4
    l=5
    i,j,k,l,m,a=genDataFunc(i,j,k,l)

elif GenericImp == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=genDataFunc(i,j,k,l)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=receivingFunc(i,j,k,l,m,a)

elif receiving == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=receivingFunc(i,j,k,l,m,a)


T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=validationFunc(i,j,k,l,m,a)

elif validation == 0:
    F=1
   
elif (T == 0 and F == 0):
    i,j,k,l,m,a=validationFunc(i,j,k,l,m,a)


T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=icFlattenFunc(i,j,k,l,m,a)

elif ICFlatten == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=icFlattenFunc(i,j,k,l,m,a)


T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=icLatestFunc(i,j,k,l,m,a)
    
elif ICLatest == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=icLatestFunc(i,j,k,l,m,a)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=dpTeradataFunc(i,j,k,l,m,a,DtLogic)

elif DPTeradata == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=dpTeradataFunc(i,j,k,l,m,a,DtLogic)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 0 and DPCam == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=dpCamFunc(i,j,k,l,m,a,DtLogic)

elif DPCam == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=dpCamFunc(i,j,k,l,m,a,DtLogic)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 0 and DPCam == 0 and OpRaw == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=opRawFunc(i,j,k,l,m,a)

elif OpRaw == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=opRawFunc(i,j,k,l,m,a)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 0 and DPCam == 0 and OpRaw == 0 and DPRefDelta == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=dpRefDelFunc(i,j,k,l,m,a)

elif DPRefDelta == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=dpRefDelFunc(i,j,k,l,m,a)

T=0
F=0
if (RawDataImp == 0 and GenericImp == 0 and receiving == 0 and validation == 0 and ICFlatten == 0 and ICLatest == 0 and DPTeradata == 0 and DPCam == 0 and OpRaw == 0 and DPRefDelta == 0 and DPRefLatest == 1):
    T=1
    i=2
    j=3
    k=4
    l=5
    m=6
    a=1
    i,j,k,l,m,a=dpRefLtstFunc(i,j,k,l,m,a)

elif DPRefLatest == 0:
    F=1

elif (T == 0 and F == 0):
    i,j,k,l,m,a=dpRefLtstFunc(i,j,k,l,m,a)

workbook.close()
