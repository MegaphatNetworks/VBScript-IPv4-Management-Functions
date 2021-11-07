'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'			IPv4 Address Library  		    '
'		               by			    '
'			   Gabriel Polmar		    '
'			  Megaphat Networks		    '
'			  www.megaphat.info		    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ConvertIPToBinary(strIP)
'' ConvertBinIPToDecimal(strBinIP)
'' MaskLength(strMask)
'' MaskLengthToIP(intMask)
'' CalcNetworkAddress(strIP, strMask)
'' CalcBroadcastAddress(strIP, strMask)
'' InvertBin(sStr)

''Convert an IP to binary
Function ConvertIPToBinary(strIP)
    '' Converts an IP Address into Binary
    
    Dim arrOctets : arrOctets = Split(strIP, ".")
    
    Dim i
    For i = 0 to UBound(arrOctets)
        Dim intOctet : intOctet = CInt(arrOctets(i))
        Dim strBinOctet : strBinOctet = ""
        
        Dim j
        For j = 0 To 7
            If intOctet And (2^(7 - j)) Then
                strBinOctet = strBinOctet & "1"
            Else
                strBinOctet = strBinOctet & "0"
            End If
        Next 
		arrOctets(i) = strBinOctet
    Next
    
    ConvertIPToBinary = Join(arrOctets, ".")
End Function
Say "ConvertIPToBinary: 255.255.255.0 = " & ConvertIPToBinary("255.255.255.0")

''Convert a binary IP to a decimal IP
Function ConvertBinIPToDecimal(strBinIP)
    '' Convert binary form of an IP back to decimal
    
    Dim arrOctets : arrOctets = Split(strBinIP, ".")
    Dim i
    For i = 0 to UBound(arrOctets)
        Dim intOctet : intOctet = 0
        
        Dim j
        For j = 0 to 7
            Dim intBit : intBit = CInt(Mid(arrOctets(i), j + 1, 1))
            If intBit = 1 Then
                intOctet = intOctet + 2^(7 - j)
            End If
        Next
        arrOctets(i) = CStr(intOctet)
    Next
    
    ConvertBinIPToDecimal = Join(arrOctets, ".")
End Function
Say "ConvertBinIPToDecimal: 11111111.11111111.11111111.00000000 = " & ConvertBinIPToDecimal("11111111.11111111.11111111.00000000")

''Convert a subnet mask to a mask length
Function MaskLength(strMask)
    '' Converts an subnet mask into a mask length in bits
    
    Dim arrOctets : arrOctets = Split(strMask, ".")
    Dim i
    For i = 0 to UBound(arrOctets)
        Dim intOctet : intOctet = CInt(arrOctets(i))
        Dim j, intMaskLength
        For j = 0 To 7
            If intOctet And (2^(7 -j)) Then
                intMaskLength = intMaskLength + 1
            End If
        Next
    Next
    
    MaskLength = intMaskLength
End Function
Say "MaskLength: 255.255.252.0 = " & MaskLength("255.255.252.0")

''Convert a mask length to a subnet mask
Function MaskLengthToIP(intMask)
    '' Converts a mask length to the decimal format mask
    
    Dim arrOctets(3)
    Dim intFullOctets : intFullOctets = (intMask - (intMask Mod 8)) / 8
    Dim i
    For i = 0 To (intFullOctets - 1)
        arrOctets(i) = "255"
    Next
    
    Dim intPartialOctetLen : intPartialOctetLen = intMask Mod 8
    Dim j
    If intPartialOctetLen > 0 Then                                 '<adding comment here
        Dim intOctet
        For j = 0 To (intPartialOctetLen - 1)
            intOctet = intOctet + 2^(7 - j)
        Next
        arrOctets(i) = intOctet : i = i + 1
    End If
    
    For j = i To 3
        arrOctets(j) = "0"
    Next
    
    MaskLengthToIP = Join(arrOctets, ".")
End Function
Say "MaskLengthToIP: 24 = " & MaskLengthToIP(24)

''Calculate the subnet network address
Function CalcNetworkAddress(strIP, strMask)
    '' Generates the Network Address from the IP and Mask
    '' Conversion of IP and Mask to binary
    
    Dim strBinIP : strBinIP = ConvertIPToBinary(strIP)
    Dim strBinMask : strBinMask = ConvertIPToBinary(strMask)
    
    '' Bitwise AND operation (except for the dot)
    Dim i, strBinNetwork
    For i = 1 to Len(strBinIP)
        Dim strIPBit : strIPBit = Mid(strBinIP, i, 1)
        Dim strMaskBit : strMaskBit = Mid(strBinMask, i, 1)
        
        If strIPBit = "1" And strMaskBit = "1" Then
            strBinNetwork = strBinNetwork & "1"
        ElseIf strIPBit = "." Then
            strBinNetwork = strBinNetwork & strIPBit
        Else
            strBinNetwork = strBinNetwork & "0"
        End If
    Next
    
    '' Conversion of Binary IP to Decimal
    CalcNetworkAddress= ConvertBinIPToDecimal(strBinNetwork)
End Function
Say "CalcNetworkAddress: 10.54.27.1,255.255.252.0 = " & CalcNetworkAddress("10.54.27.1","255.255.252.0")

''Calculate the subnet broadcast address
Function CalcBroadcastAddress(strIP, strMask)
    '' Generates the Broadcast Address from the IP and Mask
    '' Conversion of IP and Mask to binary
    
    Dim strBinIP : strBinIP = ConvertIPToBinary(strIP)
    Dim strBinMask : strBinMask = ConvertIPToBinary(strMask)
    
    '' Set each unmasked bit to 1
    Dim i, strBinBroadcast
    For i = 1 to Len(strBinIP)
        Dim strIPBit : strIPBit = Mid(strBinIP, i, 1)
        Dim strMaskBit : strMaskBit = Mid(strBinMask, i, 1)
        
        If strIPBit = "1" Or strMaskBit = "0" Then
            strBinBroadcast = strBinBroadcast & "1"
        ElseIf strIPBit = "." Then
            strBinBroadcast = strBinBroadcast & strIPBit
        Else
            strBinBroadcast = strBinBroadcast & "0"
        End If
    Next
    
    '' Conversion of Binary IP to Decimal
    CalcBroadcastAddress = ConvertBinIPToDecimal(strBinBroadcast) 
End Function
Say "CalcBroadcastAddress: 10.54.27.1,255.255.252.0 = " & CalcBroadcastAddress("10.54.27.1","255.255.252.0")

'' Invert a binary mask for Cisco-style subnetting
function InvertBin(sStr)
    sA = split(sStr,".")
    sLine = ""
    for i = 0 to ubound(sA)
        for j = 1 to len(sA(i))
            sTemp = mid(sA(i),j,1)
            sNew = cstr(1 xor cint(sTemp))
            sLine = sLine & sNew
        next
    	sLine = sLine & "."
    next
    InvertBin = left(sLine, len(sLine)-1)
end function
Say "InvertBin: 11111111.11111111.11111111.00000000 = " & InvertBin("11111111.11111111.11111111.00000000")
Say "InvertBin for submask 255.255.255.0 = " & InvertBin(ConvertIPToBinary("255.255.255.0"))
Say "InvertBin for bitmask 24 = " & InvertBin(ConvertIPToBinary(MaskLengthToIP(24)))

'' Converts an Internet host address to an Internet dot address 
Function INET_NTOA(ip)
    ip0 = Split(ip, ".")(0)
    ip1 = Split(ip, ".")(1)
    ip2 = Split(ip, ".")(2)
    ip3 = Split(ip, ".")(3)
    urlobfs = 0
    urlobfs = ip0 * 256
    urlobfs = urlobfs + ip1
    urlobfs = urlobfs * 256
    urlobfs = urlobfs + ip2
    urlobfs = urlobfs * 256
    urlobfs = urlobfs + ip3
    INET_NTOA = urlobfs
End Function
Say INET_NTOA("192.168.1.1")

'' Support function for isValidNet
Function CIDR2IP(ip, high)
    highs = "11111111111111111111111111111111"
    lows = "00000000000000000000000000000000"
    byte0 = Dec2bin(Split(ip, ".")(0))
    byte1 = Dec2bin(Split(ip, ".")(1))
    byte2 = Dec2bin(Split(ip, ".")(2))
    byte3 = Dec2bin(Split(Split(ip, ".")(3), "/")(0))
    Mask = Split(Split(ip, ".")(3), "/")(1)
    bytes = byte0 & byte1 & byte2 & byte3
    rangelow = Left(bytes, Mask) & Right(lows, 32 - Mask)
    rangehigh = Left(bytes, Mask) & Right(highs, 32 - Mask)
    iplow = bin2ip(Left(bytes, Mask) & Right(lows, 32 - Mask))
    iphigh = bin2ip(Left(bytes, Mask) & Right(highs, 32 - Mask))
    If high Then
        CIDR2IP = iphigh
    Else
        CIDR2IP = iplow
    End If
End Function


'' Support function for CIDR2IP, converts a 32 bit binary address to a dot-notated IP address
''expecting input like 00000000000000000000000000000000
Function bin2ip(strbin)
    ip0 = C2dec(Mid(strbin, 1, 8))
    ip1 = C2dec(Mid(strbin, 9, 8))
    ip2 = C2dec(Mid(strbin, 17, 8))
    ip3 = C2dec(Mid(strbin, 25, 8))
    ''combines all of the bytes into a single string
    bin2ip = ip0 & "." & ip1 & "." & ip2 & "." & ip3 
End Function

'' Support function for bin2ip, converts binary to decimal from 00000000 through 11111111
''expecting input like 00010101
Function C2dec(strbin)
    length = Len(strbin)
    dec = 0
    For x = 1 To length
        binval = 2 ^ (length - x)
        temp = Mid(strbin, x, 1)
        If temp = "1" Then dec = dec + binval
    Next
    C2dec = dec
End Function

'' Support function for CIDR2IP, converts decimal 0 through 255 to binary
''Expecting input 0 thru 255
Function Dec2bin(dec)
    Const maxpower = 7
    Const length = 8
    bin = ""
    x = cLng(dec)
    For m = maxpower To 0 Step -1
        If x And (2 ^ m) Then
            bin = bin + "1"
        Else
            bin = bin + "0"
        End If
    Next
    Dec2bin = bin
End Function


'' Generally formats a date to any format
Function fmtDateTime(sFmt, aData)
	''USAGE: fmtDateTime("{0:yyyyMMdd}", Array(now))
    Dim g_oSB : Set g_oSB = CreateObject("System.Text.StringBuilder")
    g_oSB.AppendFormat_4 sFmt, (aData)
    fmtDateTime = g_oSB.ToString()
    g_oSB.Length = 0
End Function
Say Date & " = " & fmtDateTime("{0:yyyyMMdd}", Array(date))

'' Converts a bitmask to a Cisco wildcard mask
Function GetWildcard(sSubnet)
    sBlockMask = MaskLengthToIP(sSubnet)
	sBlockBin = Trim(ConvertIPToBinary(sBlockMask))
    sBlockInv = InvertBin(sBlockBin)
    GetWildcard = ConvertBinIPToDecimal(sBlockInv)
End Function
Say "GetWildcard: 24 = " & GetWildcard(24)

'' Verifies is an IP can exist in a specific CIDR, returns boolean
Function isValidNet(sIPA, sCIDR)
	IP_Add = INET_NTOA(sIPA)
    IP_Low = INET_NTOA(CIDR2IP(sCIDR, false)) 
    IP_High= INET_NTOA(CIDR2IP(sCIDR, True))
    If (IP_Add => IP_Low) And (IP_Add <= IP_High) Then 
		isValidNet = 1 
    Else 
		isValidNet = 0
	End If
End Function
Say "isValidNet: 192.168.1.2 192.168.1.0/24 = " & CBool(isValidNet("192.168.1.2","192.168.1.0/24"))
Say "isValidNet: 192.168.1.2 192.168.2.0/24 = " & CBool(isValidNet("192.168.1.2","192.168.2.0/24"))
Say "isValidNet: 192.168.1.2 192.168.1.0/32 = " & CBool(isValidNet("192.168.1.2","192.168.2.0/24"))

Sub Say(sStr)
	wscript.echo sStr & vbcrlf
End Sub
	
