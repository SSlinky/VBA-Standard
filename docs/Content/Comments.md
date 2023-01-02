# Comments

Comments that contradict the code are worse than no comments. Always make a priority of keeping the comments up-to-date when the code changes!

Comments should be complete sentences. The first word should be capitalised, unless it is an identifier that begins with a lower case letter (never alter the case of identifiers!).

Block comments generally consist of one or more paragraphs built out of complete sentences, with each sentence ending in a period.

English should be lingua franca unless you're sure that everyone who will read the comments speaks the language you choose. If you choose a language other than English, you should use this language consistently throughout the project.

For full-line comments, the comment flag `'` should not have white space before it and the comment should be indented to the same level as the code. Comment blocks can include their own additional indentation, e.g., the args list in a method comment.

## Block Comment

Block comments generally apply to some (or all) code that follows them, and are indented to the same level as that code. Each line of a block comment starts with a # and a single space (unless it is indented text inside the comment).

Paragraphs inside a block comment are separated by a line containing a single `'`.

A block comment can be a single line and can greatly improve general comprehension of the functionality of a method.

```vb
Public Sub WriteDocument()
'   Writes the parsed markdown to a new document.

'   Reference the markdown block tree
    Dim mkdwn As BlockContainer
    Set mkdwn = mBlockStack.Items(mBlockStack.Count)

'   Exit if there's no tree.
    If Not BlockTreeHasContent(mkdwn) Then
        Throw = Errs.LexerNothingToWriteWarning
        Exit Sub
    End If

'   Write to and style the document.
    mkdwn.WriteContent AttachedDocument
    mkdwn.StyleContent AttachedDocument
End Sub
```

## Inline Comment

Use inline comments sparingly.

An inline comment is a comment on the same line as a statement. Inline comments should be separated by at least two spaces from the statement. They should start with a `'` and a single space.

Inline comments are unnecessary and in fact distracting if they state the obvious. Don’t do this:

```vb
i = i + 1       ' Increment i by 1.
```

But sometimes this is useful:

```vb
i = i - 1       ' Decrement i to account for deleted row
```

## Method Comment

Methods should contain one or more lines that explain the functionality. For methods that take arguments or raise an exception, a variation of the Google style of comment for Python is preferred. A blank line should separate this block from code or the next comment. Simpler methods can get away with a single line.

```vb
Public Function ThrowHolyHandGrenade(numberOfTheCounting As Long) As Boolean
'   Lobs a hand grenade at thy enemies who, being naughty
'   in my sight, shall snuff it.
'
'   Args:
'       numberOfTheCounting: The number to count before throwing.
'
'   Returns:
'       True if the throw was successful and the enemies snuffed it.
'
'   Raises:
'       CountError: If the number of the counting is not 3.

'   Test for incorrect counts.
    Select Case True
        Case Is numberOfTheCounting < 3 Or numberOfTheCounting = 4:
            Throw = Errs.CountError("Thou shalt not count " & numberOfTheCounting)
            Exit Function
        Case Is > 4
            Throw = Errs.CountError(numberOfTheCounting & " is right out!")
            Exit Function
    End Select

'   Throw the hand grenade and report snuffed enemies.
    Actions.Lob("Holy Hand Grenade")
    ThrowHolyHandGrenade = True
End Function
```

## Class Header Comment

Classes should contain a header comment that names the class and provides a brief explanation as to its purpose. This comment block should appear before any class related code and after the options. It should be preceeded by a single blank line and proceeded by two blank lines. The block itself should be fenced by a line of 79 dashes (to make a row of 80 characters).

Example class:

```vb
Option Explicit
Implements IFileReader

'-------------------------------------------------------------------------------
'   Class: FileReaderHttp
'   Reads the contents of a file at an http location.
'-------------------------------------------------------------------------------


' Private Backing Store
'-------------------------------------------------------------------------------
Private mLines As List
Private mNextLine As String
```

Example interface:

```vb
Option Explicit

'-------------------------------------------------------------------------------
'   Interface: IFileReader
'   An interface that describes base file reading methods.
'   This interface exists so different text objects can be swapped out,
'   e.g. file, direct text, or http get.
'-------------------------------------------------------------------------------


Public Property Get EOF() As Boolean
'   Returns True if the entire contents has been read.
    Throw = Errs.InterfaceUsedAsObject
End Property
```

## Class Section Comment

Classes should be separated out into logical sections. At the start of each section, a section header should describe it.

Each section header should be preceeded by two blank rows. It should be proceeded by a fence of 79 dashes (to make a row of 80 characters).

Example:

```vb
' Properties
'-------------------------------------------------------------------------------
Public Property Get IsEoF() As Boolean
'   Returns True if the entire contents has been read.
    IsEoF = mLines.Count = 0
End Property


' Methods
'-------------------------------------------------------------------------------
Public Function PeekNextLine() As String
'   Returns the next line to be read without advancing the pointer.
    If IsEoF Then Throw = Errs.FileReaderEOF
    PeekNextLine = mLines.Peek
End Function

Public Function ReadNextLine() As String
'   Returns the next line to be read and advances the pointer.
    If IsEoF Then Throw = Errs.FileReaderEOF
    ReadNextLine = mLines.Pop
End Function
```
