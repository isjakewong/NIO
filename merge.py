import xlrd
import xlwt
import openpyxl

workbook0 = xlrd.open_workbook("./final_2.xlsx")
workbook1 = xlwt.Workbook(encoding="utf-8",style_compression=0)
# worksheet = workbook.sheet_by_name("same")
# workbook0 = openpyxl.load_workbook("./same_selected.xlsx")
worksheet0 = workbook0.sheet_by_name("Sheet1")
worksheet1 = workbook1.add_sheet('test', cell_overwrite_ok=True)
nrows = worksheet0.nrows
ncols = worksheet0.ncols
it = iter(range(nrows-2))
for i in it:
    print(i)
    for j in range(i+1, i+2):
        for m in range(j+1, j+2):
            if worksheet0.cell_value(i, 2) == worksheet0.cell_value(j, 2) and worksheet0.cell_value(j, 2) == worksheet0.cell_value(m, 2):
                for k in range(ncols):
                    if worksheet0.cell_value(i, k) == worksheet0.cell_value(j, k) and worksheet0.cell_value(j, k) == worksheet0.cell_value(m, k):
                        worksheet1.write(i, k, worksheet0.cell_value(i, k))
                    else:
                        worksheet1.write(i, k, str(worksheet0.cell_value(i, k))+" "+str(worksheet0.cell_value(j, k))+" "+str(worksheet0.cell_value(m, k)))
                next(it)
                next(it)
            elif worksheet0.cell_value(i, 2) == worksheet0.cell_value(j, 2) and worksheet0.cell_value(i, 2) != worksheet0.cell_value(m, 2):
                for k in range(ncols):
                    if worksheet0.cell_value(i, k) == worksheet0.cell_value(j, k):
                        worksheet1.write(i, k, worksheet0.cell_value(i, k))
                    else:
                        worksheet1.write(i, k, str(worksheet0.cell_value(i, k))+" "+str(worksheet0.cell_value(j, k)))      
                next(it)         
workbook1.save("./test02.xls")
 
