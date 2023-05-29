Attribute VB_Name = "Clear"
Sub Clear()
    Dim WS As Worksheet
    Dim i As Long
    Dim HTTPreq As Object, url As String, response As String, json As String
    Dim jsonObject As Object, item As Object
    
    
    Set WS = Worksheets("Sayfa1")
    Set HTTPreq = CreateObject("MSXML2.XMLHTTP")
    
    url = "http://localhost:3000/api/TodoIsts/getall"
    
    With HTTPreq
        .Open "Get", url, False
        .send
    End With
    
    response = HTTPreq.responseText
    json = response
   
    Set objectJson = JsonConverter.ParseJson(json)
    
    WS.Cells(2, 1).Interior.ColorIndex = 0
    WS.Cells(2, 2).Interior.ColorIndex = 0
    WS.Cells(2, 1) = ""
    WS.Cells(2, 2) = ""
    
    
    i = 3
    For Each result In objectJson("data")
        WS.Cells(i, 1) = ""
        WS.Cells(i, 2) = ""
        i = i + 1
    Next
End Sub
