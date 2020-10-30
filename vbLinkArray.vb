VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbLinkArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const RANGE_ERROR As Integer = 801
Private Const SYNTAX_ERROR As Integer = 802
Private Const CONNECT_ERROR As Integer = 803

Private Self As vbArray

Public Master As vbLinkArray
Public Slave As vbLinkArray

Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "インスタンス生成時に呼び出される"
    Set Self = New vbArray
    Set Master = Nothing
    Set Slave = Nothing
End Sub

Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "インスタンス破棄化時に呼び出される"
End Sub

Private Function IndexIsValid(ByVal Index As Integer) As Boolean
Attribute IndexIsValid.VB_Description = "指定したインデックスが有効な範囲(0 <= Index < Length)であるかを判定する"
    IndexIsValid = 0 <= Index Or Index < Self.Length
End Function

Public Function Add(ByVal Value As Variant) As vbLinkArray
Attribute Add.VB_Description = "配列の末尾に新たな値を追加する"
    Call Self.Add(Value)
    If Not Slave Is Nothing Then
        Call Slave.Add(Value)
    End If
    Set Add = Me
End Function

Public Function Clear() As vbLinkArray
Attribute Clear.VB_Description = "配列に格納されている全ての値を削除する"
    Call Self.Clear
    If Not Slave Is Nothing Then
        Call Slave.Clear
    End If
    Set Clear = Me
End Function

Public Function Concat(ByVal Value As vbLinkArray) As vbLinkArray
Attribute Concat.VB_Description = "配列に別の配列の全ての値を追加する"
    Call Self.Concat(Value)
    If Not Slave Is Nothing Then
        Call Slave.Concat(Value)
    End If
    Set Concat = Me
End Function

Public Sub Connect(ByRef Value As vbLinkArray)
Attribute Connect.VB_Description = "配列を別の配列と関連付ける"
    Dim CurrentValue As vbLinkArray
    Set CurrentValue = Value.Master
    While Not CurrentValue Is Nothing
        If CurrentValue Is Me Then
            Err.Raise CONNECT_ERROR
            Exit Sub
        End If
        Set CurrentValue = CurrentValue.Master
    Wend
    Set CurrentValue = Value.Slave
    While Not CurrentValue Is Nothing
        If CurrentValue Is Me Then
            Err.Raise CONNECT_ERROR
            Exit Sub
        End If
        Set CurrentValue = CurrentValue.Slave
    Wend
    If Value Is Me Then
        Err.Raise CONNECT_ERROR
        Exit Sub
    End If
    Set Value.Master = Me
    Set Slave = Value
End Sub

Public Function Copy(ByVal Value As vbLinkArray) As vbLinkArray
Attribute Copy.VB_Description = "配列の全ての値を、別の配列の全ての値と変更する"
    Call Self.Copy(Value)
    If Not Slave Is Nothing Then
        Call Slave.Copy(Value)
    End If
    Set Copy = Me
End Function

Public Sub Disconnect()
Attribute Fill.VB_Description = "配列の関連付けを解除する"
    Set Master = Nothing
    Set Slave = Nothing
End Sub

Public Function Every(ByVal Callback As String) As Boolean
Attribute Every.VB_Description = "配列の全ての値がコールバックを満たすか判定する"
    Every = Self.Every(Callback)
End Function

Public Function Fill(ByVal Value As Variant) As vbLinkArray
Attribute Fill.VB_Description = "配列の全ての値を引数の値に置き換える"
    Call Self.Fill(Value)
    If Not Slave Is Nothing Then
        Call Slave.Fill(Value)
    End If
    Set Fill = Me
End Function

Public Function Filter(ByVal Callback As String) As vbLinkArray
Attribute Filter.VB_Description = "配列からコールバックを満たさない値を削除する"
    Call Self.Filter(Callback)
    If Not Slave Is Nothing Then
        Call Slave.Filter(Callback)
    End If
    Set Filter = Me
End Function

Public Function Find(ByVal Callback As String) As Variant
Attribute Find.VB_Description = "配列の中でコールバックを満たす一番前の値を返す"
    Find = Self.Find(Callback)
End Function

Public Function FindLast(ByVal Callback As String) As Variant
Attribute FindLast.VB_Description = "配列の中でコールバックを満たす一番後ろの値を返す"
    FindLast = Self.FindLast(Callback)
End Function

Public Function FindIndex(ByVal Callback As String) As Integer
Attribute FindIndex.VB_Description = "配列の中でコールバックを満たす一番前のインデックスを返す"
    FindIndex = Self.FindIndex(Callback)
End Function

Public Function FindLastIndex(ByVal Callback As String) As Integer
Attribute FindLastIndex.VB_Description = "配列の中でコールバックを満たす一番後ろのインデックスを返す"
    FindLastIndex = Self.FindLastIndex(Callback)
End Function

Public Sub ForEach(ByVal Callback As String)
Attribute ForEach.VB_Description = "配列の全ての値でコールバックを実行する"
    Call Self.ForEach(Callback)
    If Not Slave Is Nothing Then
        Call Slave.ForEach(Callback)
    End If
End Sub

Public Sub FromString(ByVal Value As String)
Attribute FromString.VB_Description = "文字列を配列に変換する"
    Call Self.FromString(Value)
    If Not Slave Is Nothing Then
        Call Slave.FromString(Value)
    End If
End Sub

Public Sub FromRange(ByVal Value As Range)
Attribute FromString.VB_Description = "セル範囲を配列に変換する"
    Call Self.FromString(Value)
    If Not Slave Is Nothing Then
        Call Slave.FromRange(Value)
    End If
End Sub

Public Function Includes(ByVal Value As Variant) As Boolean
Attribute Includes.VB_Description = "配列に値が格納されているかを判定する"
    Includes = Self.Includes(Value)
End Function

Public Function IndexOf(ByVal Value As Variant) As Integer
Attribute IndexOf.VB_Description = "配列の中で引数と等しい値の一番前のインデックスを返す"
    IndexOf = Self.IndexOf(Value)
End Function

Public Function LastIndexOf(ByVal Value As Variant) As Integer
Attribute LastIndexOf.VB_Description = "配列の中で引数と等しい値の一番後ろのインデックスを返す"
    LastIndexOf = Self.LastIndexOf(Value)
End Function

Public Function Map(ByVal Callback As String) As vbLinkArray
Attribute Map.VB_Description = "配列の全ての値をコールバックの返り値に置き換える"
    Call Self.Map(Callback)
    If Not Slave Is Nothing Then
        Call Slave.Map(Callback)
    End If
    Set Map = Me
End Function

Public Function Reduce(ByVal Callback As String, _
        Optional ByVal InitialValue As Variant = 0) As Variant
Attribute Reduce.VB_Description = "配列の全ての値でコールバックを実行し、その返り値を結合する"
    Reduce = Self.Reduce(Callback, InitialValue)
End Function

Public Function Remove(ByVal Index As Integer) As vbLinkArray
Attribute Remove.VB_Description = "配列から引数のインデックスの値を削除する"
    Call Self.Remove(Index)
    If Not Slave Is Nothing Then
        Call Slave.Remove(Index)
    End If
    Set Remove = Me
End Function

Public Function Reverse() As vbLinkArray
Attribute Reverse.VB_Description = "配列の並び順を逆転する"
    Set Reverse = Self.Reverse()
End Function

Public Function Some(ByVal Callback As String) As Boolean
Attribute SoVB_Description = "配列の1つ以上の値がコールバックを満たすか判定する"
    Some = Self.Some(Callback)
End Function

Public Function Sort() As vbLinkArray
Attribute Sort.VB_Description = "配列を小さい順に並び替える"
    Call Self.Sort
    If Not Slave Is Nothing Then
        Call Slave.Sort
    End If
    Set Sort = Me
End Function

Public Function ToString() As String
Attribute ToString.VB_Description = "配列を文字列に変換する"
    ToString = Self.ToString
End Function

Public Property Get Value(ByVal Index As Integer) As Variant
Attribute Value.VB_Description = "指定したインデックスの値を取得する"
    Value = Self(Index)
End Property

Public Property Let Value(ByVal Index As Integer, _
        ByVal Value As Variant)
Attribute Value.VB_Description = "指定したインデックスの値を変更する"
    Self(Index) = Value
End Property
