Function f(a, b)
  Dim aNum, bNum
  ' Explicitly convert to numbers to avoid implicit type issues
  aNum = CDbl(a)
  bNum = CDbl(b)
  If aNum > bNum Then
    MsgBox "a is greater than b"
  ElseIf aNum < bNum Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end function