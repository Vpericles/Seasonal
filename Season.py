from statsmodels.tsa.seasonal import seasonal_decompose
import xlrd as x1
import matplotlib.pyplot as plt
import xlsxwriter

# Reads Excel with a single time series
# no dates column, top row with series name
file = "IPCAalone.xls"
wb = x1.open_workbook(file)
s1 = wb.sheet_by_index(0)
r = s1.nrows
print("dataset of",r-1,"observations")
series = []
i = 1
while i < r:
    series.append(s1.cell_value(i,0))
    i = i+1
# print(series)
result = seasonal_decompose(series, model='additive', period = 12)
ssfactors = result.seasonal
l = len(ssfactors)
print("Dataset of",l,"seasonal factors")
workbook = xlsxwriter.Workbook("SeasonFactors.xlsx")
worksheet = workbook.add_worksheet()
runs = 1
worksheet.write(0,0,"Season Factors")
while runs <= l:
    worksheet.write(runs,0,ssfactors[runs-1])
    runs = runs+1
workbook.close()

# print(result.trend)
# print(result.seasonal)
# print(result.resid)
# print(result.observed)
result.plot()
plt.show()