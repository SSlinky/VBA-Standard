# Comments

Comments that contradict the code are worse than no comments. Always make a priority of keeping the comments up-to-date when the code changes!

Comments should be complete sentences. The first word should be capitalised, unless it is an identifier that begins with a lower case letter (never alter the case of identifiers!).

Block comments generally consist of one or more paragraphs built out of complete sentences, with each sentence ending in a period.

English should be lingua franca unless you're sure that everyone who will read the comments speaks the language you choose. If you choose a language other than English, you should use this language consistently throughout the project.

## Block Comments

Block comments generally apply to some (or all) code that follows them, and are indented to the same level as that code. Each line of a block comment starts with a # and a single space (unless it is indented text inside the comment).

Paragraphs inside a block comment are separated by a line containing a single `'`.

## Inline Comments

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

Inline comments can also appear above the line or block of lines they relate to. This can improve general comprehension of the functionality of a method.

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
