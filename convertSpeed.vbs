If WScript.Arguments.Count > 0 Then
    Dim speedValue
    speedValue = CDbl(WScript.Arguments(0))
    Dim convertedSpeed
    Dim unit

    If speedValue > 1000 Then
        convertedSpeed = speedValue / 1000
        unit = "Gbps"
    Else
        convertedSpeed = speedValue
        unit = "Mbps"
    End If

    convertedSpeed = FormatNumber(convertedSpeed, 0)
    WScript.Echo convertedSpeed & " " & unit
End If
