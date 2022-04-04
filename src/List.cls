VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Public Name As String
Private mParameter() As Variant
Private mValue() As Variant
Private mGP() As Variant
Private mNoMatch As Boolean
Private mSize As Long
Private mRow As Long
Private mBOL As Boolean
Private mEOL As Boolean

Public Property Get Parameter(Optional Row As Variant) As Variant
    If mSize > 0 And Not mBOL And Not mEOL Then
        If Not IsMissing(Row) Then
            If Row >= 0 And Row <= mSize - 1 Then
                mRow = Row
            End If
        End If
        Parameter = mParameter(mRow)
    End If
End Property

Public Property Get Value(Optional Row As Variant) As Variant
    If mSize > 0 Then
        If Not IsMissing(Row) Then
            If Row >= 0 And Row <= mSize - 1 Then
                mRow = Row
            End If
        End If
        Value = mValue(mRow)
    End If
End Property

Public Property Get GeneralPurpose(Optional Row As Variant) As Variant
    If mSize > 0 Then
        If Not IsMissing(Row) Then
            If Row >= 0 And Row <= mSize - 1 Then
                mRow = Row
            End If
        End If
        Value = mGP(mRow)
    End If
End Property

Public Property Get Size() As Long
    Size = mSize
End Property

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Get EOL() As Boolean
    EOL = mEOL
End Property

Public Property Get BOL() As Boolean
    BOL = mBOL
End Property

Private Property Get NoMatch() As Boolean
    NoMatch = mNoMatch
End Property

Public Function MoveFirst()
    If mSize > 0 Then
        mRow = 0
        mBOL = False
        mEOL = False
    End If
End Function

Public Function MoveNext()
    mRow = mRow + 1
    If mRow >= mSize Then
        mRow = mSize
        mBOL = False
        mEOL = True
    Else
        mBOL = False
        mEOL = False
    End If
End Function

Public Function MovePrevious()
    mRow = mRow - 1
    If mRow < 0 Then
        mRow = -1
        mBOL = True
        mEOL = True
    Else
        mBOL = False
        mEOL = False
    End If
End Function

Private Sub Class_Initialize()
    mRow = -1
    mSize = 0
    mBOL = True
    mEOL = True
End Sub

Public Function Add(Paramenter As Variant, Optional Value As Variant, Optional GeneralPurpose As Variant)
    Dim Position As Long

    If Paramenter <> "" Then
        mRow = mSize
        ReDim Preserve mParameter(mRow)
        ReDim Preserve mValue(mRow)
        ReDim Preserve mGP(mRow)
        
        mParameter(mRow) = Paramenter
        mValue(mRow) = Value
        mGP(mRow) = GeneralPurpose
        
        mSize = mSize + 1
        mBOL = False
        mEOL = False
    End If
End Function

Public Function InsertAt(Index As Long, Paramenter As Variant, Optional Value As Variant, Optional GeneralPurpose As Variant)
    Dim Position As Long
    
    If Paramenter <> "" Then
        mRow = mSize
        ReDim Preserve mParameter(mRow)
        ReDim Preserve mValue(mRow)
        ReDim Preserve mGP(mRow)
        
        For Position = mSize To Index + 1 Step -1
            mParameter(Position) = mParameter(Position - 1)
            mValue(Position) = mValue(Position - 1)
            mGP(Position) = mGP(Position - 1)
        Next
        
        mParameter(Position) = Paramenter
        mValue(Position) = Value
        mGP(Position) = GeneralPurpose
        
        mSize = mSize + 1
        mBOL = False
        mEOL = False
    End If
End Function

Public Function RemoveAt(Index As Long)
    Dim Position As Long
    If mSize <= 0 Or Index < 0 Or Index >= mSize Then Exit Function
    If Size > 1 Then
        For Position = Index + 1 To mSize - 1
            mParameter(Position - 1) = mParameter(Position)
            mValue(Position - 1) = mValue(Position)
            mGP(Position - 1) = mGP(Position)
        Next
        mRow = mSize - 1
        ReDim Preserve mParameter(mRow - 1)
        ReDim Preserve mValue(mRow - 1)
        ReDim Preserve mGP(mRow - 1)
    End If
    mSize = mSize - 1
    If Size <= 0 Then
        mBOL = True
        mEOL = True
    Else
        mBOL = False
        mEOL = False
    End If
End Function

Public Function GetAttribute(ByRef Parameter As String, ByRef Value As String, Optional ByRef GeneralPurpose As Variant, Optional Row As Variant)
    Parameter = ""
    Value = ""
    If mSize > 0 Then
        If Not IsMissing(Row) Then
            If Row >= 0 And Row <= mSize - 1 Then
                mRow = Row
            End If
        End If
        Parameter = mParameter(mRow)
        Value = mValue(mRow)
        GeneralPurpose = mGP(mRow)
    End If
End Function
Public Function IndexOf(Value As String, Optional Index As Variant = Null) As Long
    Dim Position As Long
    mNoMatch = True
    If mSize > 0 Then
        If IsNull(Index) Or Row >= mSize Then
            Position = 0
        Else
            Position = Index
        End If
        For Position = Position To mSize - 1
            If mValue(Position) = Value Then
                mRow = Position
                IndexOf = Position
                mNoMatch = False
                Exit Function
            End If
        Next
    End If
    IndexOf = -1
End Function

Public Function Contains(Value As String, Optional Index As Variant = Null) As Boolean
    Dim Position As Long
    mNoMatch = True
    If mSize > 0 Then
        If IsNull(Index) Or Row >= mSize Then
            Position = 0
        Else
            Position = Index
        End If
        For Position = Position To mSize - 1
            If mValue(Position) = Value Then
                mRow = Position
                Contains = True
                mNoMatch = False
                Exit Function
            End If
        Next
    End If
    IndexOf = -1
End Function
       
