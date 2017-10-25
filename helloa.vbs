dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://raw.githubusercontent.com/doubleshot43/go/master/test.txt", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    '.write xHttp.responseBody
    '.savetofile "c:\temp\test.txt", 2 '//overwrite
end with