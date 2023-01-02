# Naming Conventions

Naming conventions are an important aspect of writing clean and maintainable code. In Visual Basic for Applications (VBA), naming conventions help to ensure that your code is easy to read and understand, and that it follows best practices for naming variables, functions, and other code elements.

It's important to note that you may sometimes find yourself working with legacy code that does not follow a consistent naming convention. In these situations, you may need to work outside of your preferred naming convention in order to understand and modify the existing code.

However, when writing new code, it's important to follow a consistent naming convention in order to make your code as readable and maintainable as possible. By following a clear set of rules for naming your code elements, you can make it easier for other developers to understand your code and make changes to it if needed.

## Syle Conventions

There are a lot of different naming styles. It helps to be able to recognize what naming style is being used, independently from what they are used for.

The following naming styles are distinguished in this guide:

* `mixedCase` also known as lower CamelCase.
* `CapitalizedWords` also known as CapWords, CapCase, PascalCase, or upper CamelCase.

     Note: When using acronyms in CapWords, capitalize all the letters of the acronym. Thus HTTPServerError is better than HttpServerError.
* `UPPER_CASE_WITH_UNDERSCORES`

## Names to Avoid

Never use the characters 'l' (lower case el), 'O' (upper case oh), or ‘I’ (uppercase letter eye) as single character variable names. These characters can be difficult to distinguish from others.

Hungarian notation should not be used. This is the practice of prefixing the variable type onto the variable name. It serves to increase variable length and complexity while providing no benefit. In addition, if you were to change a type, say from Integer to Long, you'd have to find and rename the variable wherever it is used.

Avoid names that conflict with core functionality or other names in the project. This will avoid bugs that arise from the compiler selecting one or another version when not fully qualified.

## ASCII Compatibility

All identifiers must use ASCII-only identifiers, and should use English words wherever feasible (in many cases, abbreviations and technical terms are used which aren’t English).

## Meaningful Names

Names should be descriptive enough to convey meaning. The use of abbreviations, acronyms, and single letter names should be avoided unless commonly understood, e.g., `var`, `i`, `j`. Abbreviations used should be consistent throughout the project.

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
