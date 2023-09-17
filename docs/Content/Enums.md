# Enums

Enumerated types are a way of setting strict categories that the programmer can assign but not change. Enums help to make the code self documenting when used in favour of constants.

Enums can be declared with explicit or implicit values. Explicit assignment should only be used in cases where it matters what the values are, e.g.:

* Mutually inclusive enums.
* External processes dictating the values.
* Values expected to be used in calculations.

## Implicit Enums

Enums that have implicit value assignment are used when only one category or option can be assigned. In the below example, a file can only be of one type, the file types aren't expectected to be involved in any calculations, and it does not matter in which order they are listed. The majority of enums can be written in this way.

```vb
Public Enum FileTypes
    ExcelWorkbook
    PortableDocumentFormat
    PowerpointDeck
    TextDocument
    WordDocument
End Enum
```

## Explicit Enumns

Explicit enums are written where the value matters and cannot simply be a zero based, single increment. This can occur when your values are mutually inclusive, are influenced by processes outside your control, or the values are expected to be used in calculations.

**Mutually inclusive enums.**

```vb
Public Enum ProgramOptions
    CanWrite = 1            ' 2^0
    PrintWarnings = 2       ' 2^1
    HaltOnWarnings = 4      ' 2^2
    LogToFile = 8           ' 2^3
    TestModeEnabled = 16    ' 2^4
End Enum
```

These options are not mutually exclusive. Setting their values as exponents of 2 allows us to perform bitwise operations to check for "hot" or "flagged" bits in the same variable.

**External processes dictating the values.**

Your organisation may use specific error values so that components can be used together without conflicting and for the standardisation of log writing and reading.

```vb
Public Enum ErrorCode
    MightBeBroken = vbObjectError + MYPROGNUM + 1
    ContactYourAdministrator = vbObjectError + MYPROGNUM + 2
    EverythingOnFire = vbObjectError + MYPROGNUM + 3
End Enum
```

The use of your programme number in a constant also allows for this to change without requiring a refactor.

**Values expected to be used in calculations.**

Uncommonly, values may be used in calculations. Therefore, it may simplify code to use the enum directly rather than a `Select` statement.

```vb
Public Enum ScoreWeighting
    Low = 20
    Medium = 50
    High = 80
    Critical = 120
End Enum
```

This way, a score may be weighted by its criticality, e.g., `weightedScore = score * scoreWeight \ 100`.
