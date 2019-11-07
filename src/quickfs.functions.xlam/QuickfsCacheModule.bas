Attribute VB_Name = "QuickfsCacheModule"
Option Explicit
Option Private Module

Private CachedValues As New Dictionary
Private CachedTimestamp As New Dictionary
Private RecachedValues As New Dictionary
Private RecachedTimestamp As New Dictionary
Private Recaching As Boolean

Public Function ClearCache()
    CachedValues.RemoveAll
    CachedTimestamp.RemoveAll
    RecachedValues.RemoveAll
    RecachedTimestamp.RemoveAll
    Recaching = False
End Function

Public Function StartRecache()
    Recaching = True
End Function

Public Function StopRecache()
    Recaching = False
    Dim key As Variant
    For Each key In RecachedValues.keys
        If TypeName(RecachedValues.Item(key)) = "Collection" Or TypeName(RecachedValues.Item(key)) = "ErrorPoint" Then
            Set CachedValues.Item(key) = RecachedValues.Item(key)
        Else
            CachedValues.Item(key) = RecachedValues.Item(key)
        End If
        CachedTimestamp.Item(key) = RecachedTimestamp.Item(key)
    Next
    RecachedValues.RemoveAll
    RecachedTimestamp.RemoveAll
End Function

Public Function IsCached(ByVal key As String, Optional skip As Boolean = False) As Boolean
    ' Return boolean true if cached value for key is available within the cache timeout
    IsCached = False
    
    If skip Then
        Exit Function
    End If
    
    If Not Recaching Then
        If CachedTimestamp.Exists(key) Then
            If CachedTimestamp.Item(key) + (CACHE_TIMEOUT_MINUTES / 60 / 24) >= Now() Then IsCached = True
        End If
    Else
        If RecachedTimestamp.Exists(key) Then
            If RecachedTimestamp.Item(key) + (CACHE_TIMEOUT_MINUTES / 60 / 24) >= Now() Then IsCached = True
        End If
    End If
End Function

Public Sub SetCachedValue(ByVal key As String, ByVal dataValue As Variant)
    ' Set cached value and timestamp for key
    If Not Recaching Then
        If TypeName(dataValue) = "Collection" Or TypeName(dataValue) = "ErrorPoint" Then
            Set CachedValues.Item(key) = dataValue
        Else
            CachedValues.Item(key) = dataValue
        End If
        CachedTimestamp.Item(key) = Now()
    Else
        If TypeName(dataValue) = "Collection" Or TypeName(dataValue) = "ErrorPoint" Then
            Set RecachedValues.Item(key) = dataValue
        Else
            RecachedValues.Item(key) = dataValue
        End If
        RecachedTimestamp.Item(key) = Now()
    End If
End Sub

Public Function GetCachedValue(ByVal key As String) As Variant
    ' Retrieve cached value for key
    If Not Recaching Then
        If CachedValues.Exists(key) Then
            If TypeName(CachedValues.Item(key)) = "Collection" Or TypeName(CachedValues.Item(key)) = "ErrorPoint" Then
                Set GetCachedValue = CachedValues.Item(key)
            Else
                GetCachedValue = CachedValues.Item(key)
            End If
        Else
            GetCachedValue = CVErr(xlErrNA) ' return #NA
        End If
    Else
        If RecachedValues.Exists(key) Then
            If TypeName(RecachedValues.Item(key)) = "Collection" Or TypeName(RecachedValues.Item(key)) = "ErrorPoint" Then
                Set GetCachedValue = RecachedValues.Item(key)
            Else
                GetCachedValue = RecachedValues.Item(key)
            End If
        Else
            GetCachedValue = CVErr(xlErrNA) ' return #NA
        End If
    End If
End Function

Public Function IsCachedError(key As String) As Boolean
    IsCachedError = (TypeName(GetCachedValue(key)) = "ErrorPoint")
End Function

Public Function CachedToQFS(key As String, Optional index As Integer)
    If TypeName(GetCachedValue(key)) = "Collection" Then
        Dim list As Collection
        Set list = GetCachedValue(key)
        If TypeName(index) = "Empty" Or index = 0 Then
            CachedToQFS = CollectionToString(list)
        ElseIf list.count < index Then
            CachedToQFS = CVErr(xlErrNull)
        Else
            CachedToQFS = list(index)
        End If
    ElseIf IsCachedError(key) Then
        Dim ep As ErrorPoint
        Set ep = GetCachedValue(key)
        Err.Raise ep.code, ep.name, ep.description
    Else
        CachedToQFS = GetCachedValue(key)
    End If
End Function
