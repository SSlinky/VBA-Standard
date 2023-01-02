# Global Variables

Global variables are variables that are declared outside of any specific procedure or function, and are therefore available to be used by any procedure or function in the code.

Another particular quirk of VBA is that sometimes these variables will be cleared when the programme exits and sometimes they will not. This can cause dirty data between runs and unwanted results.

## Avoid Them

Avoid using global variables whenever possible: Global variables can make it harder to understand and maintain your code, as they can be accessed and modified from any part of the code. Instead of using global variables, consider how the code can be designed to use local variables.

## Store Together

Store all global variables you cannot easily avoid in a module created just for that purpose. This can help to make your code easier to understand and maintain.

The same module could be a home for global constants.

## Use a Naming Convention

To make it easier to identify global variables in your code, consider using a naming convention that sets them apart from other variables. For example, you could use a prefix or suffix to indicate that a variable is global.
