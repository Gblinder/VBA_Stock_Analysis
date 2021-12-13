# VBA_Stock_Analysis
VBA Stock Analysis Challenge
Gabi Blinder
gblinder@gmail.com
Analysis of Code Refactoring

#  Advantages and Disadvantages of Refactoring Code
Refactoring, or altering code for efficacy, is best done in small intentional steps. There are advantages and disadvantages to this process:
## Advantages:
* Errors in logistics are easily recognized easily in structure code that contains nested conditionals and loops.
* In our case, using Excel shows program logic more simply, not tied to the order that the underlying code is written.
* VBA interpretation (Excel) of code can show patterns that are not easy to see in the source code.
* 
## Disadvantages:
* A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
* A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
* A complex unstructured code is usually best to split in several functions.
* Refactoring process can affect the testing outcomes.

Applying these pros and cons to refactoring the original VBA script?
Code refactoring is changing or updating code but not affecting software’s functionality or external behavior of the application. If ion the future, you intend to update or troubleshoot your code, if it is complicated or hard to understand, there ensues a foreseeable issue.
The refactoring process is a way of “cleaning up”code, making it easier for you or future users of your code or program. Ideally making it easier to understand, change, and improving upon efficacy.
