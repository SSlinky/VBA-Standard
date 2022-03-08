# Indentation

Indentation helps the reader understand the programme flow by clearly delineating logical blocks of code. A single level of indentation is four spaces.

## Module Level Statements and Method Signatures

Module level statements and method signatures are the lowest level of indentation. These inclde key words such as `Option`, `Dim`, `Public`, `Private`, `Const`, `Property`, `Function`, and `Sub`.

## Logical Code Blocks

Code that opens or closes a block is at the same level of indentation as the preceeding and proceeding code. Code within the block has one higher level of indentation. This means that entering a block will always indent once and exiting will outdent once. The only exception to this rule is exiting a select case block (see below).

**The wrong way to do it.**

```vb
Public Sub ContrivedExample()
Dim i As Long, j As Long

For i = 1 to 100
For j = 1 to 100
If j < i Then
If i Mod j = 0 Then
Debug.Print j & " divisible by " & i
End If
End If
Next j
Next i
End Sub
```

**The right way to do it.**

```vb
Public Sub ContrivedExample()
    Dim i As Long, j As Long

    For i = 1 to 100
        For j = 1 to 100
            If j < i Then
                If i Mod j = 0 Then
                    Debug.Print i & " divisible by " & j
                End If
            End If
        Next j
    Next i
End Sub
```

## Close Open Statements

Some keywords both close and open a block, e.g. `Else`, `ElseIf`. These should be at the same level of indentation as the code that opens and closes that block.

## Case Statements

`End Select` outdents twice since the one statement closes both the last `Case` and the entire `Select Case` block.

```vb
Select Case turtleAge
    Case < 8
        MsgBox "Young turtle!"
    Case < 50
        MsgBox "Not a baby!"
    Case Else
        MsgBox "Old turtle!"
End Select
```

## Line Continuation

Line continuation should be indented or aligned vertically to clearly show that they are a part of the previous line.

Vertical alignment should be used when the arguments begin on the originating line. Otherwise, a hanging indent should be used.

**An example of vertical alignment.**

```vb
Public Property Get FormattedTime(Optional x As Long) As String
'   Formats the passed in tick count or the current timer
'   tick count as a human-readable time.

    Dim t As Long ' time
    Dim h As Long ' hours
    Dim m As Long ' minutes
    Dim s As Long ' seconds
    Dim d As Long ' decimals
    
    t = IIf(x > 0, x, Me.ElapsedTime)
    
    s = CLng(t / 1000) Mod 60
    m = CLng(t / 60000) Mod 60
    h = CLng(t / 3600000)
    d = CLng((t / 1000 - Int(t / 1000)) * 1000)
    
    FormattedTime = Format(h, "00") & ":" _
                    & Format(m, "00") & ":" _
                    & Format(s, "00") & "." _
                    & Format(d, "000")
    
End Property
```

**A long method signature with a hanging indent.**

```vb
Function MyLongFunctionSignature( _
    firstArgument As Long, _
    secondArgument As Long, _
    thirdArgumentSomeObject As Turtle
)
```
