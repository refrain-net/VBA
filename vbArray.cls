VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const RANGE_ERROR As Integer = 801
Private Const SYNTAX_ERROR As Integer = 802

Private Self() As Variant

Public Length As Integer

Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "インスタンス生成時に呼び出される"
    Length = 0
    ReDim Self(Length)
End Sub

Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "インスタンス破棄化時に呼び出される"
End Sub

Private Function IndexIsValid(ByVal Index As Integer) As Boolean
Attribute IndexIsValid.VB_Description = "指定したインデックスが有効な範囲(0 <= Index < Length)であるかを判定する"
    IndexIsValid = 0 <= Index Or Index < Length
End Function

Public Function Add(ByVal Value As Variant) As vbArray
Attribute Add.VB_Description = "配列の末尾に新たな値を追加する"
    ReDim Preserve Self(Length)
    Self(Length) = Value
    Length = Length + 1
    Set Add = Me
End Function

Public Function Clear(Optional ByVal OnlyMaster As Boolean = False) As vbArray
Attribute Clear.VB_Description = "配列に格納されている全ての値を削除する"
    Call Class_Initialize
    Set Clear = Me
End Function

Public Function Concat(ByVal Value As vbArray) As vbArray
Attribute Concat.VB_Description = "配列に別の配列の全ての値を追加する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Value.Length
        Call Add(Value(CurrentIndex))
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Concat = Me
End Function

Public Function Copy(ByVal Value As vbArray) As vbArray
Attribute Copy.VB_Description = "配列の全ての値を、別の配列の全ての値と変更する"
    Call Clear
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Value.Length
        Call Add(Value(CurrentIndex))
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Copy = Me
End Function

Public Function Every(ByVal Callback As String) As Boolean
Attribute Every.VB_Description = "配列の全ての値がコールバックを満たすか判定する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Not Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            Every = False
            Exit Function
        End If
        CurrentIndex = CurrentIndex + 1
    Wend
    Every = True
End Function

Public Function Fill(ByVal Value As Variant) As vbArray
Attribute Fill.VB_Description = "配列の全ての値を任意の値に置き換える"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        Self(CurrentIndex) = Value
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Fill = Me
End Function

Public Function Filter(ByVal Callback As String) As vbArray
Attribute Filter.VB_Description = "配列からコールバックを満たさない値を削除する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Not Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            Call Remove(CurrentIndex)
        Else
            CurrentIndex = CurrentIndex + 1
        End If
    Wend
    Set Filter = Me
End Function

Public Function Find(ByVal Callback As String) As Variant
Attribute Find.VB_Description = "配列の中でコールバックを満たす一番前の値を返す"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            Find = Self(CurrentIndex)
            Exit Function
        End If
        CurrentIndex = CurrentIndex + 1
    Wend
    Find = Empty
End Function

Public Function FindLast(ByVal Callback As String) As Variant
Attribute FindLast.VB_Description = "配列の中でコールバックを満たす一番後ろの値を返す"
    Dim CurrentIndex As Integer
    CurrentIndex = Length - 1
    While CurrentIndex > -1
        If Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            FindLast = Self(CurrentIndex)
            Exit Function
        End If
        CurrentIndex = CurrentIndex - 1
    Wend
    FindLast = Empty
End Function

Public Function FindIndex(ByVal Callback As String) As Integer
Attribute FindIndex.VB_Description = "配列の中でコールバックを満たす一番前のインデックスを返す"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            FindIndex = CurrentIndex
            Exit Function
        End If
        CurrentIndex = CurrentIndex + 1
    Wend
    FindIndex = -1
End Function

Public Function FindLastIndex(ByVal Callback As String) As Integer
Attribute FindLastIndex.VB_Description = "配列の中でコールバックを満たす一番後ろのインデックスを返す"
    Dim CurrentIndex As Integer
    CurrentIndex = Length - 1
    While CurrentIndex > -1
        If Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            FindLastIndex = CurrentIndex
            Exit Function
        End If
        CurrentIndex = CurrentIndex - 1
    Wend
    FindLastIndex = -1
End Function

Public Sub ForEach(ByVal Callback As String)
Attribute ForEach.VB_Description = "配列の全ての値でコールバックを実行する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        Call Run(Callback, Self(CurrentIndex), CurrentIndex, Me)
        CurrentIndex = CurrentIndex + 1
    Wend
End Sub

Public Sub FromRange(ByVal Value As Range)
Attribute FromString.VB_Description = "セル範囲を配列に変換する"
    Call Clear
    Dim CurrentValue As Variant
    For Each CurrentValue In Value
        Call Add(CurrentValue.Value)
    Next CurrentValue
End Sub

Public Function Includes(ByVal Value As Variant) As Boolean
Attribute Includes.VB_Description = "配列に値が格納されているかを判定する"
    Includes = IndexOf(Value) > -1
End Function

Public Function IndexOf(ByVal Value As Variant) As Integer
Attribute IndexOf.VB_Description = "配列の中で引数と等しい値の一番前のインデックスを返す"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Self(CurrentIndex) = Value Then
            IndexOf = CurrentIndex
            Exit Function
        End If
        CurrentIndex = CurrentIndex + 1
    Wend
    IndexOf = -1
End Function

Public Function LastIndexOf(ByVal Value As Variant) As Integer
Attribute LastIndexOf.VB_Description = "配列の中で引数と等しい値の一番後ろのインデックスを返す"
    Dim CurrentIndex As Integer
    CurrentIndex = Length - 1
    While CurrentIndex > -1
        If Self(CurrentIndex) = Value Then
            LastIndexOf = CurrentIndex
            Exit Function
        End If
        CurrentIndex = CurrentIndex - 1
    Wend
    LastIndexOf = -1
End Function

Public Function Map(ByVal Callback As String) As vbArray
Attribute Map.VB_Description = "配列の全ての値をコールバックの返り値に置き換える"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        Self(CurrentIndex) = Run(Callback, Self(CurrentIndex), CurrentIndex, Me)
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Map = Me
End Function

Public Function Parse(ParamArray Values() As Variant) As vbArray
Attribute Reduce.VB_Description = "配列を空にし、引数の値を全て追加する"
    Dim CurrentValue As Variant
    For Each CurrentValue In Values
        Call Add(CurrentValue)
    Next CurrentValue
    Set Parse = Me
End Function

Public Function Reduce(ByVal Callback As String, _
        Optional ByVal InitialValue As Variant = 0) As Variant
Attribute Reduce.VB_Description = "配列の全ての値でコールバックを実行し、その返り値を結合する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        InitialValue = Run(Callback, InitialValue, Self(CurrentIndex), CurrentIndex, Me)
        CurrentIndex = CurrentIndex + 1
    Wend
    Reduce = InitialValue
End Function

Public Function Remove(ByVal Index As Integer) As vbArray
Attribute Remove.VB_Description = "配列から引数のインデックスの値を削除する"
    If Not IndexIsValid(Index) Then
        Err.Raise RANGE_ERROR
        Exit Function
    End If
    While Index < Length - 1
        Self(Index) = Self(Index + 1)
        Index = Index + 1
    Wend
    Length = Length - 1
    ReDim Preserve Self(Length - 1)
    Set Remove = Me
End Function

Public Function Reverse() As vbArray
Attribute Reverse.VB_Description = "配列の並び順を逆転する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    Dim TempIndex As Integer
    Dim CurrentValue As Variant
    While CurrentIndex < (Length - Length Mod 2) / 2
        TempIndex = Length - CurrentIndex - 1
        CurrentValue = Self(CurrentIndex)
        Self(CurrentIndex) = Self(TempIndex)
        Self(TempIndex) = CurrentValue
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Reverse = Me
End Function

Public Function Some(ByVal Callback As String) As Boolean
Attribute SoVB_Description = "配列の1つ以上の値がコールバックを満たすか判定する"
    Dim CurrentIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        If Run(Callback, Self(CurrentIndex), CurrentIndex, Me) Then
            Some = True
            Exit Function
        End If
        CurrentIndex = CurrentIndex + 1
    Wend
    Some = False
End Function

Public Function Sort() As vbArray
Attribute Sort.VB_Description = "配列を小さい順に並び替える"
    Dim CurrentValue As Variant
    Dim CurrentIndex As Integer
    Dim NextValue As Variant
    Dim NextIndex As Integer
    CurrentIndex = 0
    While CurrentIndex < Length
        NextIndex = Length - 1
        While NextIndex > CurrentIndex
            CurrentValue = Self(CurrentIndex)
            NextValue = Self(NextIndex)
            If CurrentValue > NextValue Then
                Self(CurrentIndex) = NextValue
                Self(NextIndex) = CurrentValue
            End If
            NextIndex = NextIndex - 1
        Wend
        CurrentIndex = CurrentIndex + 1
    Wend
    Set Sort = Me
End Function

Public Function ToString() As String
Attribute ToString.VB_Description = "配列を文字列に変換する"
    If Length > 0 Then
        ToString = "[ " & Join(Self, ", ") & " ]"
    Else
        ToString = "[ ]"
    End If
End Function

Public Property Get Value(Optional ByVal Index As Integer = -1) As Variant
Attribute Value.VB_Description = "指定したインデックスの値を変更する"
Attribute Value.VB_UserMemId = 0
    If Index = -1 Then
        Value = Self
        Exit Property
    ElseIf Not IndexIsValid(Index) Then
        Err.Raise RANGE_ERROR
        Exit Property
    End If
    Value = Self(Index)
End Property

Public Property Let Value(ByVal Index As Integer, _
        ByVal Value As Variant)
Attribute Value.VB_Description = "指定したインデックスの値を変更する"
    If Not IndexIsValid(Index) Then
        Err.Raise RANGE_ERROR
        Exit Property
    End If
    Self(Index) = Value
End Property
