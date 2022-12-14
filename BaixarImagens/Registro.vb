Public Class Registro
    Private _numNota As String
    Private _marca As String
    Private _link As String
    Public Sub New(ByVal NumNota As String, ByVal Marca As String, ByVal Link As String)
        Me.NumNota = NumNota
        Me.Marca = Marca
        Me._link = Link
    End Sub

    Public Property NumNota() As String
        Get
            Return _numNota
        End Get
        Set(ByVal Valor As String)
            _numNota = Valor
        End Set
    End Property

    Public Property Marca() As String
        Get
            Return _marca
        End Get
        Set(ByVal Valor As String)
            _marca = Valor
        End Set
    End Property

    Public Property Link() As String
        Get
            Return _link
        End Get
        Set(ByVal Valor As String)
            _link = Valor
        End Set
    End Property
End Class
