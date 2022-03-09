# Line Breaks At Operators

Traditional programming convention was to break after the operator. Mathemeticians and their publishers follow the convention of breaking before the operator. This enhances readability over the traditional way to break after the operator.

Following the tradition from mathematics usually results in more readable code.

**The wrong way to do it.**

```vb
income = grossWages + _
         superannuation + _
         dividends + _
         (capitalGain - capitalLoss) - _
         garnishments
```

**The right way to do it.**

```vb
income = grossWages _
         + superannuation _
         + dividends _
         + (capitalGain - capitalLoss) _
         - garnishments
```
