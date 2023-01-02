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
