Option Explicit

Sub PrintOptions(strFpSetting)
    ' Create DOM object
    Dim objDOM, objGP, objNodes
    Set objDOM = CreateObject("Msxml2.DOMDocument.6.0")
    objDOM.Load strFpSetting
    For Each objGP In objDOM.SelectNodes("/Root/GrepPatterns/GrepPattern")
        WScript.Echo objGP.getAttribute("index") & ": " & objGP.getAttribute("name")
    Next
    Set objGP = Nothing
    Set objDOM = Nothing
End Sub

Sub Main()
    Dim strFpSetting: strFpSetting = WScript.Arguments(0)
    WScript.Echo "PrintOptions.vbs: strFpSetting=" & strFpSetting
    PrintOptions(strFpSetting)
End Sub
Main