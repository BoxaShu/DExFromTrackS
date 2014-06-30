Imports System.Xml.Serialization

Public Class Cage
    <XmlAttribute> _
    Public Name As String ' = "1"
    <XmlElement("Al")> _
    Public Along_List As New List(Of Along)
    'Добавить сравнение по имени
    Public Overrides Function GetHashCode() As Integer
        'Xor has the advantage of not overflowing the integer.
        Return Name.GetHashCode
    End Function
    Public Overloads Overrides Function Equals(ByVal Obj As Object) As Boolean
        Dim oKey As Section = CType(Obj, Section)
        Return (oKey.Name = Me.Name)
    End Function
End Class