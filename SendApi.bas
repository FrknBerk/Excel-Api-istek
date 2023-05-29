Attribute VB_Name = "SendApi"
Sub SenAPIrequest()
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
    
    WS.Cells(2, 1).Interior.ColorIndex = 5
    WS.Cells(2, 2).Interior.ColorIndex = 6
    WS.Cells(2, 1) = "TITLE"
    WS.Cells(2, 2) = "DESCRIPTION"
    
    
    i = 3
    For Each result In objectJson("data")
        WS.Cells(i, 1) = result("title")
        WS.Cells(i, 2) = result("description")
        i = i + 1
    Next
   
    
End Sub

