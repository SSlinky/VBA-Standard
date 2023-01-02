# Long Over Integer

**Integers** provide no memory or performance benefit over using the **Long** data type. This [SackOverflow](https://stackoverflow.com/questions/26409117/why-use-integer-instead-of-long) post goes into depth as to why.

From an old [MSDN article](http://msdn.microsoft.com/en-us/library/office/aa164506%28v=office.10%29.aspx):

> Traditionally, VBA programmers have used integers to hold small numbers, because they required less memory. In recent versions, however, VBA converts all integer values to type Long, even if they're declared as type Integer. So there's no longer a performance advantage to using Integer variables; in fact, Long variables may be slightly faster because VBA does not have to convert them.

The edge case to using Integers as a preference is when interacting with some legacy external components outside your control. From the above linked SO post:

> So, in summary, there's almost no good reason to use an Integer type these days. Unless you need to Interop with an old API call that expects a 16 bit int, or you are working with large arrays of small integers and memory is at a premium.
