Class Car
    Public Color
    Public Brand
    Private InnerModel
    Private InnerEngine

    '值类型属性
    Public Property Get Model()
        Model=InnerModel
    End Property
    Public Property Let Model(Value)
        InnerModel=Value
    End Property

    '对象类型属性
    Public Property Set Engine(Value)
        Set InnerEngine=Value
    End Property
    Public Property Get Engine()
        Set Engine=InnerEngine
    End Property

    Private Sub Class_Initialize()
    'This event is called when an instance of the class is instantiated
    'Initialize properties here and perform other start-up tasks
        Color="red"
    End Sub

    Private Sub Class_Terminate()
    'This event is called when a class instance is destroyed
    'either explicitly (Set objClassInstance = Nothing) or
    'implicitly (it goes out of scope)
    End Sub

    Public Sub Go()
        MsgBox(Color + " "+ InnerModel +" car are go.")
    End Sub
End Class

Set myCar = New Car
myCar.Model="8888"
myCar.go