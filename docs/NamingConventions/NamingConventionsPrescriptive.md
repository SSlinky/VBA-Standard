# Prescriptive Naming Conventions

Names should be descriptive enough to convey meaning. The use of abbreviations, acronyms, and single letter names should be avoided unless commonly understood, e.g., `var`, `i`, `j`.

## Object Names

Object names, e.g., Classes, Forms, Modules, Sheets (in Excel) etc. should use the CapWords convention.

Interfaces should be prefixed with an I and then follow CapWords, e.g., `ITextReader`.

## Variables

Public variables should follow the CapCase convention. Local or private variables should follow the mixed case convention.

Member variables, e.g., backing store for properties, should be prefixed with an `m`. Modern convention typically favours a leading underscore but this is illegal in VBA so we fall back to the older member style.

Some commonly understood variables don't need to be 'descriptive' because their intent is already widely understood. That being said, they should only be used within the scope of that understanding. It's okay to `Dim i As Long` explicitly when `i` is to be used as an iterable in a loop.

**The wrong way to do it.**

```vb
Dim i As Long
Dim j As Long

i = Rows.Count
For j = 1 To i
    ...
Next j
```

**The right way to do it.**

```vb
Dim i As Long
Dim maxRows As Long

maxRows = Rows.Count
For i = 1 To maxRows
    ...
Next i
```

## Constants

Constants are usually defined on a module level and written in all capital letters with underscores separating words. Examples include `MAX_OVERFLOW` and `TOTAL`.

## Methods

Methods should follow the CapCase convention. They should not include an underscore unless they are an interface implementation (concrete) or event handler method. Event handlers may need to follow mixed case as they are bound by the name of the object they are subscribed to. Examples: `ITextReader_ReadLine()` or `mStreamReader_OnReadCharacter`.

Method arguments should follow the local variable style.

## Events

Event methods should include an adverb and a verb that indicates the action that triggered them and when. Examples: `Public Event BeforeSave()` or `Public Event AfterSave()`. This makes it clear exactly where in the process the event is fired.

```vb
Public Event BeforeSave()
Public Event AfterSave()

Public Sub Save()
    RaiseEvent BeforeSave()
    ...
    RaiseEvent AfterSave()
End Sub
```

Event handlers are a special type of method. They follow the naming convention of `objectSubscribedTo_EventMethodName`. They should follow the standard method convention as far as it is possible while adhering to this rule. Following on with the Event example:

```vb
Private WithEvents mMyObject As MyClass

Public Sub mMyObject_AfterSave()
    ...
End Sub
```

## Types

Types should follow the CapCase convention. As types are used in fairly narrow scope, it can be helpful to suffix the name with `Type` to differentiate it from the variable, Example:

```vb
Type CustomerType
  Name As String
  Address As String
  Phone As String
End Type

Sub Example()
    Dim customer As CustomerType
    customer.Name = "John Smith"
    customer.Address = "123 Main St."
    customer.Phone = "0118 999 881 999 119 725 3"

    Debug.Print customer.Name & " lives at " & _
                customer.Address & " and can be reached at " & _
                customer.Phone
End Sub
```
