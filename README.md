# Excel-Network-Ping-Monitor

# Excel Network Ping Monitor

A simple **Excel VBA tool** that allows users to monitor the status of multiple IP addresses directly inside Microsoft Excel.
The script pings each IP address and updates the **status and response time** in the spreadsheet, including **color-coded indicators** for quick monitoring.

## Features

* Ping multiple IP addresses from Excel
* Display **Online / Offline** status
* Show **response time (ms)**
* Automatic **color-coded status**
* Simple **one-button execution**
* Useful for **network monitoring, device checks, and quick diagnostics**

## Status Color Indicators

| Status        | Response Time | Color  |
| ------------- | ------------- | ------ |
| Online (Fast) | ≤ 50 ms       | Green  |
| Online (Slow) | 51–150 ms     | Yellow |
| Very Slow     | >150 ms       | Orange |
| Offline       | No response   | Red    |

## Example Excel Layout

| Device | IP Address  | Status         |
| ------ | ----------- | -------------- |
| Router | 192.168.1.1 | Online (2 ms)  |
| DNS    | 8.8.8.8     | Online (30 ms) |
| Server | 10.10.10.1  | Offline        |

## Setup Instructions

1. Open **Microsoft Excel**.
2. Press **ALT + F11** to open the VBA editor.
3. Click **Insert → Module**.
4. Paste the VBA code below.
5. Save the file as **.xlsm (Macro Enabled Workbook)**.
6. Add a button from **Developer → Insert → Button (Form Control)**.
7. Assign the macro **PingAllIPs** to the button.

## VBA Code

```vba
Sub PingAllIPs()

Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim ip As String
Dim objPing As Object
Dim objStatus As Object
Dim response As Long

Set ws = ActiveSheet
lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

For i = 2 To lastRow

    ip = ws.Cells(i, 3).Value
    
    If ip <> "" Then
    
        Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}") _
        .ExecQuery("select * from Win32_PingStatus where address='" & ip & "'")

        For Each objStatus In objPing
        
            If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            
                ws.Cells(i, 4).Value = "Offline"
                ws.Cells(i, 4).Interior.Color = RGB(255, 0, 0)
                
            Else
            
                response = objStatus.ResponseTime
                ws.Cells(i, 4).Value = "Online (" & response & " ms)"
                
                If response <= 50 Then
                    ws.Cells(i, 4).Interior.Color = RGB(0, 176, 80)
                ElseIf response <= 150 Then
                    ws.Cells(i, 4).Interior.Color = RGB(255, 255, 0)
                Else
                    ws.Cells(i, 4).Interior.Color = RGB(255, 192, 0)
                End If
                
            End If
            
        Next objStatus
        
    End If
    
Next i

End Sub
```

## Customization

You can easily change the columns used in the script.

Example:

```vba
ip = ws.Cells(i, 3).Value
```

* `3` = Column C (IP address)

```vba
ws.Cells(i, 4).Value
```

* `4` = Column D (Status output)

## Use Cases

* Network device monitoring
* Lab environment testing
* Infrastructure status checks
* IT troubleshooting
* Device availability reports

## Requirements

* Microsoft Excel with **VBA support**
* Windows OS (uses WMI `Win32_PingStatus`)

## License

MIT License
