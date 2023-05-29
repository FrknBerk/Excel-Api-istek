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
    For Each pokemon In objectJson("data")
        WS.Cells(i, 1) = pokemon("title")
        WS.Cells(i, 2) = pokemon("description")
        i = i + 1
    Next
   
    
End Sub

