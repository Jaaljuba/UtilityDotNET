Public Class EntityClass

End Class

Public Class GenericList
    Private _id As String
    Private _value As String

    Public Property Id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            _id = value
        End Set
    End Property

    Public Property Value() As String
        Get
            Return _value
        End Get
        Set(ByVal value As String)
            _value = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal id As String, ByVal value As String)
        _id = id
        _value = value
    End Sub
End Class
