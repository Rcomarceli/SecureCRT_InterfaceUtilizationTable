#$language = "VBScript"
#$interface = "1.0"
' This program is currently confirmed to be working on version 6.5.4 of SecureCRT, a SSH client that allows for scripting via VbScript.
' crt.screen commands below interact with the SecureCRT client, mostly telling it to input keys or wait for certain output before continuing.
' Hostnames and Hostname structures have been substituted for security purposes.

' The objective of the program was to:
' - Mostly automate a NOC task where technicians needed to gather transmit and receive utilization percentages for certain hosts and their interfaces for a daily report
' - Be compatiable with *every* NOC engineer's environment without the need to install a coding language or program, as all users lacked install permissions.

' For compatiability, I coded this to work with SecureCRT and output the data in a formatted table via printf in a linux environment, as this was the common environment for every NOC engineer. 

' For implementation, the program logs into devices, grabs transmit and receive utilizations for certain interfaces, processes that data into percentages, and displays that data in a readable table. More specifically, it:
' - SSHs over a list of defined hostname and interface pairs
' - Grabs transmit and receive via Regex
' - Converts that data into percentages
' - Outputs a printf command that when entered on a Linux server, displays a formatted table of data that looks like below:

' ```
' Site1     TB Router 1     TB Router 1     TB Router 2     TB Router 2     Edge Router 1 
' Int:      Te1/0/5         Te1/0/3         Te1/0/5         Te1/0/3         Te0/1/0 
' Tx:       0.4%            13.7%           3.9%            11.8%           11.8% 
' Rx:       0.4%            11%             3.5%            0.4%            3.5%
 
' Site2     TB Router 1     TB Router 1     TB Router 2     TB Router 2     Edge Router 1 
' Int:      Te1/0/2         Te1/0/3         Te1/0/2         Te1/0/3         Te0/1/0 
' Tx:       2%              0.4%            0.4%            1.2%            5.9% 
' Rx:       1.2%            0.4%            0.4%            2.4%            2.4%
 
' Site3     TB Router 1 
' Int:      Gi0/0/5 
' Tx:       0.8% 
' Rx:       1.2%
' ```
' This specific table format was chosen to mimic the Onenote table these data values would later be manually input to.


crt.Screen.Synchronous = True


' User Settings
Dim passwd : passwd = "" 'If empty, prompts user for password.
const decimalPlaces = 1 'desired amount of decimal places for percentages
const authTimeout = 10 'seconds
const homeDirectoryStr = "~"
const columnWidth = 10 'width of table column. if too low, we get a "space" error




Sub Main

    ' grab password if empty 
    If (passwd = "") Then
        passwd = crt.Dialog.Prompt("Enter Pass for Login." & vbcrlf & vbcrlf & "Note: Edit the script with your password to have" & vbcrlf & "it entered automatically." , "Password Input", "", True)
        If (passwd = "") Then 
            ' exit script if cancel is pressed/no password is input
            Exit Sub
        End If
    End If

    ' clear current text with ctrl + u
    crt.Screen.SendKeys "^(u)"

    ' accounts for cursor not being on "home directory" row
    crt.Screen.Send(vbcr)
    crt.Screen.WaitForString(homeDirectoryStr)

    ' these are the device & interface pairs that the program interates through. Progmatically, this could have been implemented with much less redundancy via a 2 dimensional array, but the goal was to facilitate adjustments by other engineers with little coding experience in the event that devices/interfaces needed to be added or removed. 
    
    Dim hostnameInterfacePairs(10)
    hostnameInterfacePairs(0) = "Site1_TBRouter1 Te1/0/5"
    hostnameInterfacePairs(1) = "Site1_TBRouter1 Te1/0/3"
    hostnameInterfacePairs(2) = "Site1_TBRouter2 Te1/0/5"
    hostnameInterfacePairs(3) = "Site1_TBRouter2 Te1/0/3"
    hostnameInterfacePairs(4) = "Site1_EdgeRouter1 Te0/1/0"
    hostnameInterfacePairs(5) = "Site2_TBRouter1 Te1/0/2"
    hostnameInterfacePairs(6) = "Site2_TBRouter1 Te1/0/3"
    hostnameInterfacePairs(7) = "Site2_TBRouter2 Te1/0/2"
    hostnameInterfacePairs(8) = "Site2_TBRouter2 Te1/0/3"
    hostnameInterfacePairs(9) = "Site2_EdgeRouter3 Te0/1/0" 
    hostnameInterfacePairs(10) = "Site3_TBRouter1 Gi0/0/5"

    ' Iterate through the interface pairs, grabbing specific interface utilizations and processing them into percentages
    Dim amtOfPairs : amtOfPairs = UBound(hostnameInterfacePairs)
    Dim columns
    ReDim columns(amtOfPairs, 3)
    For i = 0 to amtOfPairs
        Dim arr
        arr = calcHandler(hostnameInterfacePairs(i))

        If (isArray(arr) = False) Then
            If (arr = 2 OR arr = 3) Then
                ' if returnCode is authfail or unhandled error, exit script
                Exit Sub
            End If
        End If

        For j = 0 to 3
        columns(i, j) = arr(j)
        Next
    Next

    ' Manually log out after grabbing the last interface's data
    crt.Screen.Send("logout" & vbcr)
    crt.Screen.waitForString(homeDirectoryStr)

    ' Begin the eventual printf statement that will be used to display the table on the Linux server
    crt.Screen.Send("printf '\n\n\n")

    ' Generate rest of printf string, starting with a row for the first 4 columns and starting another row every 4 columns.
    renderRow columns, 0, 4
    crt.Screen.Send("\n ")
    renderRow columns, 5, 9
    crt.Screen.Send("\n ")
    renderRow columns, 10, 10
    
    ' End the printf statement with some spaces to differentiate the table from the rest of the client output and send it via carriage return.
    crt.Screen.Send("\n \n\n\n\n'" & vbcr)


End Sub


' main functions

' main handler function that calls all the other ones
Function calcHandler(hostnameInterfacePair)
    ' grab hostname and interface from string
    Dim arrPair, hostname, interface
    arrPair = Split(hostnameInterfacePair)
    hostname = arrPair(0)
    interface = arrPair(1)

    ' if hostname doesn't match, login to next device
    Dim currentLine : currentLine = getCurrentLine()
    Dim returnCode, currentDevice

    If InStr(currentLine, homeDirectoryStr) Then
        returnCode = loginToDevice(hostname, passwd, AuthTimeout)
    ElseIf InStr(currentLine, hostname) Then
        ' do nothing
    ElseIf Len(currentLine) = 0 Then
        crt.Screen.waitForString("")
    Else
        crt.Screen.Send("logout" & vbcr) 
        crt.Screen.waitForString(homeDirectoryStr)
        returnCode = loginToDevice(hostname, passwd, AuthTimeout)
    End If

    If (returnCode = 2 OR returnCode = 3) Then
        ' password failed, end script
        calcHandler = returnCode
        Exit Function
    ElseIf (returnCode = 1) Then
        calcHandler = Array(hostname, interface, "timeout", "timeout")
        Exit Function
    End If


    ' calculate txrx and get array
    Dim TxRx, resultArr, deviceName
    TxRx = calcTxRx(interface)

    resultArr = Array(hostname, interface, TxRx(0), TxRx(1))
    resultArr(2) = resultArr(2) & "%%"
    resultArr(3) = resultArr(3) & "%%"

    calcHandler = resultArr

    ' if this command isn't here, currentLine is an empty string => device will logout and log back in even if the next interface is on the same device.
    crt.Screen.waitForString(hostname)
End Function

' returns a return code. 0 = ok, 1 = timeout, 2 = authfail, 3 = weird error
' need to return the code to a variable or else we get weird "can't call sub" errors
Function loginToDevice(hostname, passwd, AuthTimeout)
    crt.Screen.Send "ssh " & hostname & vbcr

    ' array of possible shell prompts upon trying to ssh into network device
    Dim shellPrompts(2)
    shellPrompts(0) = "sword:"
    shellPrompts(1) = "authentication failed"
    shellPrompts(2) = "#"

    Do 
        Dim output
        output = crt.Screen.WaitForStrings(shellPrompts, AuthTimeout)
        Select Case output
            Case 0
                crt.Dialog.MessageBox "Auth. timed out for " & hostname & "!"
                crt.Screen.SendKeys "^(c)"
                crt.Screen.waitForString(homeDirectoryStr)
                loginToDevice = 1
                Exit Function
            Case 1 ' "password:"
                crt.Screen.Send passwd & vbcr
            Case 2 ' "authentication failed"
                crt.Dialog.MessageBox "Auth. Failed! Check if password is correct. Exiting Script."
                crt.Screen.SendKeys "^(c)"
                loginToDevice = 2
                Exit Function
            Case 3 ' "auth success"
                loginToDevice = 0
                Exit Do
            Case Else
                crt.Dialog.MessageBox "Unhandled Error! Exiting Script."
                loginToDevice = 3
                Exit Function
            End Select
    Loop
End Function


' returns array(tx, rx) for Tx and Rx percentages
Function calcTxRx(interface) 
    ' fetch reliability, txload, rxload string
    Dim txloadRxloadStr
    crt.Screen.Send "sh int " & interface & " | i relia" & vbcr
    crt.Screen.WaitForString vbcr 
    txloadRxloadStr = crt.Screen.ReadString(vbCrlf)
    txloadRxloadStr = Replace(txloadRxloadStr, vbLf, "")

    ' regex used to match the digit / digit "pattern"
    Dim re, matches
    Set re = New regexp
    With re
        .Pattern ="(\d+)\/(\d+)"
        .Global = True 
    End With

    ' expected output: matches = Array("255/255", "tx/255", "rx/255")
    Set matches = re.Execute(txloadRxloadStr)

    Dim txArr, rxArr, tx, rx, calcTxRxArr

    txArr = Split(matches(1), "/") ' "tx/255" --> ["tx", 255]
    rxArr = Split(matches(2), "/") ' "rx/255" --> ["rx", 255]
    tx = txArr(0) ' "tx"
    rx = rxArr(0) ' "rx"

    ' now that we've isolated tx and rx, we can calculate percentages
    tx = Round(tx / 255 * 100, decimalPlaces)
    rx = Round(rx / 255 * 100, decimalPlaces)

    calcTxRxArr = Array(tx, rx)
    calcTxRx = calcTxRxArr
End Function

' renders a table row using columns
Function renderRow(columnArray, lowerBound, upperBound)
    Dim headerRow, site
    Dim table

    site = Left(columnArray(lowerBound, 0), 4)
    headerRow = Array(site, "Int:", "Tx:", "Rx:")
    ' trimming header device names Site1_TBRouter1 --> TBRouter_01, etc
    isolateDeviceNames columnArray, lowerBound, upperBound

    ' insert header column as the first column
    table = insertColumn(columnArray, headerRow, lowerBound, upperBound)

    ' output to send to printf function for table
    Dim output

    For j = 0 to 3 
        output = output + " \n "

        For i = lowerBound to upperBound + 1
            output = output + leftJustified(table(i,j), columnWidth)
        Next
        
        ' get rid of excess space on last column
        output = Rtrim(output)
    Next

    crt.Screen.Send(output)
End Function


' utility functions

' ensure columnWidth > length of column value. if not, ==> 'Space' runtime error
Function leftJustified(ColumnValue, ColumnWidth)
    ' account for extra width from percent sign
    If InStr(ColumnValue, "%%") then
        ColumnWidth = ColumnWidth + 1
    End If

    leftJustified =  ColumnValue & Space(ColumnWidth - Len(ColumnValue))
End Function


Function getCurrentLine()
    Dim nRow, nColumn, currentLine
    nRow = crt.Screen.CurrentRow
    nColumn = crt.Screen.CurrentColumn - 1
    getCurrentLine = crt.Screen.Get(nRow, 0, nRow, nColumn)
End Function

' table is Arr(4,3), arr is Arr(0,3), lowerbound and upperbound is int
' inserts a column into an existing table. used to insert the header column
Function insertColumn(table, arr, lowerBound, upperBound)
  Dim i, output

  ReDim output(upperBound + 1, 3)
    For i = upperBound + 1 To lowerBound + 1 Step -1
        For j = 0 to 3
            output(i,j) = table(i - 1, j)
        Next
    Next

    For j = 0 to 3
        output(lowerBound, j) = arr(j)
    Next
        insertColumn = output
End Function

' changes hostNames like "Site#_Zone#_DeviceType_DeviceNumber" to "DeviceType_DeviceNumber" for readability
Sub isolateDeviceNames(table, lowerBound, upperBound)
    For i = lowerBound To upperBound
        table(i, 0) = Right(table(i, 0), 5)
    Next
End Sub




' fixed column fuction from https://stackoverflow.com/questions/795568/how-to-make-the-columns-in-vbscript-fixed
' insert array utility from https://developer.rhino3d.com/guides/rhinoscript/array-utilities/
