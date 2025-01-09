Function CheckAndAccess(obj, propertyName)
  If IsObject(obj) And Not IsNull(obj) Then
    If TypeName(obj) = "Object" Then
      'Check for specific object type if needed
    End If
    If obj.HasProperty(propertyName) Then
      result = obj(propertyName)
    Else
      result = "Property not found."
    End If
  Else
    result = "Object is null or not an object."
  End If
  CheckAndAccess = result
End Function

Dim myObject
Set myObject = CreateObject("Scripting.Dictionary")

'Safe access:
Dim value : value = CheckAndAccess(myObject, "Item")
MsgBox value 'Output: Property not found.

Set myObject = Nothing
value = CheckAndAccess(myObject, "Item")
MsgBox value 'Output: Object is null or not an object.