# Indentation

Indentation helps the reader understand the programme flow by clearly delineating logical blocks of code. A single level of indentation is four spaces.

1. Module level statements and method signatures are the lowest level of indentation. These inclde key words such as `Option`, `Dim`, `Public`, `Private`, `Const`, `Property`, `Function`, and `Sub`.
2. Code that opens or closes a block is at the same level of indentation as the preceeding and proceeding code. Code within the block has one higher level of indentation. This means that entering a block will always increase the indentation by one and exiting will decrease by one. The only exception to this rule is exiting a case block.
3. Some keywords both close and open a block, e.g. `Else`, `ElseIf`. These should be at the same level of indentation as the code that opens and closes that block.

## The wrong way to do it

```vb
Public Sub ContrivedExample()
Dim i As Long, j As Long

For i = 1 to 100
For j = 1 to 100
If i < 50 Then
If i % j Then
Debug.Print i & " divisible by " & j
End If
End If
Next j
Next i
End Sub
```

## The right way to do it

```vb
Public Sub ContrivedExample()
    Dim i As Long, j As Long

    For i = 1 to 100
        For j = 1 to 100
            If i < 50 Then
                If i % j Then
                    Debug.Print i & " divisible by " & j
                End If
            End If
        Next j
    Next i
End Sub
```

## Case Statements

Case statements are conditional checks that form the closure of the previous and the opening of the next blocks. This is similar to how `Else` works. What makes a case statement special is `End Select` forms the closure of the last case and the entire switch. They therefore reduce code indentation by two instead of one.

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
