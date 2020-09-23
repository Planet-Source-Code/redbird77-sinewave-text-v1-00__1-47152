VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFastString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' file    : CFastString.cls
' revised : 2001 April 30
' project : pSineWaveText.vbp
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77

Option Explicit

' TODO: Add .BufferSize property.
Private Const BUFFER_SIZE = 1000000

Private m_sBuffer As String
Private m_lGood   As Long

Public Sub Append(ByRef AppendString As String)
    
    Dim iLen As Integer
    
    'On Error GoTo ErrorHandler
    
    iLen = Len(AppendString)
     
    ' If starting position will be out of bounds then ...
    ' TODO: If m_lGood + iLen > Len(m_sBuffer) Then
    If m_lGood + 1 > Len(m_sBuffer) Then

        ' TODO: Do not assume AppendString to be < BUFFER_SIZE.
        m_sBuffer = m_sBuffer & String$(BUFFER_SIZE, vbNullChar)
        
        Debug.Print "Allocating...New length of buffer: " & Len(m_sBuffer)

    End If
     
    ' Mid starting position runs from 1 to length of string.
     
    ' Change a selected portion of m_sBuffer.
    Mid$(m_sBuffer, m_lGood + 1, iLen) = AppendString
    
    ' Update the starting position.
    m_lGood = m_lGood + iLen
        
    Exit Sub
    
'ErrorHandler:
'
'    If Err.Number = 5 Then
'        m_sBuffer = m_sBuffer & String$(BUFFER_SIZE, vbNullChar)
'        Debug.Print "Allocating...New length of buffer: " & Len(m_sBuffer)
'    End If
'
'    Resume 'Next
End Sub

Private Sub Class_Initialize()
    
    ' Initialize buffer.
    m_sBuffer = String$(BUFFER_SIZE, vbNullChar)
    
    m_lGood = 0
    
End Sub

Public Property Get Buffer() As String
    
    Buffer = Left$(m_sBuffer, m_lGood)
    
End Property

Public Sub Clear()

    m_sBuffer = String$(BUFFER_SIZE, vbNullChar)
    m_lGood = 0
    
End Sub
