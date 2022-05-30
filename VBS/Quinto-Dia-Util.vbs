Dim s
Dim j
Dim D

'D = cDate("05" & "/" & Month(Now()) & "/" & Year(Now()))

    For x = 1 To 10
        j = x + 4 & "/" & Month(Now()) & "/" & Year(Now())
        s = Weekday(j)
		
            If s <> 1 then				
                Exit For
            End If
			
    Next
	
D = CDate(j)

if Weekday(D) = 7 Then
	D = D+1
end if
if Weekday(D) <> 6 Then
	D = D+1
end if

msgbox D
