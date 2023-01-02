# Comments

Comments are an important part of any programming project, as they can help to explain the purpose and functionality of the code to other developers, or to future maintainers of the code.

Comments can be polarising with some people adamant that everything should be commented, while other purists think good code should be self-documenting, i.e., so readable that it doesn't require code.

The reality is that the best way to comment is somewhere in the middle.

## Do

- Aim to write code that is self-documenting. Well named classes, variables, and methods can go a long way to reducing the need for explanations. As can well designed code, e.g., following SOLID principles (as far as they relate to VBA).
- Use comments to explain a method at a glance. What it does, what its arguments are for, and what it returns.
- Use comments to explain logical sections of code. This reduces the mental load when scanning for a specific line.

## Don't

- Overcomment. The right amount of commnets really add value but past that they start to add obscurity and can be more difficult to maintain.
- Explain things that are obvious.
- Update your code without reviewing the comments. Comments that are incorrect or out of date are worse than no comments at all.

## Where to Comment

- Class explanation at the top (below options).
- Class section headers that logically separate out code into properties, methods, helpers, etc.
- Methods inside the signature
