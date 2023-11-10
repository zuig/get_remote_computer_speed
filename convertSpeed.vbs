If WScript.Arguments.Count > 0 Then
    Dim speedValue
    speedValue = CDbl(WScript.Arguments(0))
    Dim convertedSpeed
    Dim unit

    If speedValue >= 1000000000 Then
        convertedSpeed = speedValue / 1000000000
        unit = "Gbps"
    Else
        convertedSpeed = speedValue / 1000000
        unit = "Mbps"
    End If

    If Int(convertedSpeed) = convertedSpeed Then
        convertedSpeed = CStr(Int(convertedSpeed))
    Else
        convertedSpeed = FormatNumber(convertedSpeed, 1)
    End If

    WScript.Echo convertedSpeed & " " & unit
End If
