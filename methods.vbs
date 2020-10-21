'无参无返
Sub Fn1
    MsgBox("hello")
END SUB

'有参无返
Sub Fn2(a)

End Sub

'无参有返
Function F3()
    F3=3
End Function

'有参有返
Function F4(x)
    F4=4
End Function


Fn1()
Fn2(2)
Dim r3,r4
r3=F3()
r4=F4(4)