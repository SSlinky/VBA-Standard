# Less Indentation

While indentation is necessary to improve readability, the recommendation is to _require_ as little as possible of it.

This can be achieved by checking negative or base cases and exiting early rather than executing code within the condition block.

**The wrong way to do it.**

```vb
Function Fibonacci(n As Long, Optional nums As Collection) As Collection
'   Initialise nums if it's the first recursion.
    If nums Is Nothing Then
        Set nums = New Collection
    End If
    
    If n > 0 Then
        Dim cnt As Long
        cnt = nums.Count

'       Add the next number in the sequence.
        Select Case cnt
            Case 0
                nums.Add 0
            Case 1
                nums.Add 1
            Case Else
                nums.Add nums(cnt) _
                         + nums(cnt - 1)
        End Select

'       Recursive call to add the next iteration.
        Set Fibonacci = Fibonacci(n - 1, nums)

'   I've forgotten the condition.        
'   This else is too far away from the if.
    Else    
        Set Fibonacci = nums
        Exit Function
    End If

End Function
```

**The right way to do it.**

You can see we are spending less time indented, the conditional blocks are smaller, and as a result the function is easier to read.

```vb
Function Fibonacci(n As Long, Optional nums As Collection) As Collection
'   Initialise nums if it's the first recursion.
    If nums Is Nothing Then
        Set nums = New Collection
    End If
    
    
'   Base condition to exit when 0 or less passed in.
    If n <= 0 Then
        Set Fibonacci = nums
        Exit Function
    End If
    
    Dim cnt As Long
    cnt = nums.Count
    
'   Add the next number in the sequence.
    Select Case cnt
        Case 0
            nums.Add 0
        Case 1
            nums.Add 1
        Case Else
            nums.Add nums(cnt) _
                     + nums(cnt - 1)
    End Select
    
'   Recursive call to add the next iteration.
    Set Fibonacci = Fibonacci(n - 1, nums)

End Function
```
