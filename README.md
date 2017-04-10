# CodeAnalyst
A development tool that helps FoxPro developers identify areas of code that should or could be refactored.

Code Analyst is extensible, allowing developers to create their own refactoring rules and then enable or disable them as needed. Some of the code was based on the Code References tool.

Rules can be associated with different aspects of code. For example, an Object rule might analyse all objects on a form to ensure they are using a naming convention.

You can select to analyze an individual file, the current project or a directory (and subdirectories).

#### Default Refactoring Rules

There are four different types of rules:
* File
* Method/Function
* Object
* Line

Note that each rule may be disabled if you don't want to use it.

#### Default Method Rules

* Check if there is an unreasonable ratio of comments to code in a method
* Check to see if there is a RETURN within a WITH statement, a known cause of C5 errors
* Check to see if there are too many lines in a method (defaults to 150) - which is a sign of readability
* Check to see if you have more than a certain number of loop structures in code (defaults to 5)
* Verifies that each function has a RETURN value
* Verifies that a function doesn't have more than 3 return values

#### Default Object Rules

* Checks if a button named Cancel does not have the Cancel property set to .T.
* Checks if you are simply using the default object names
* Checks if there are methods in the object that have similar number of lines that could therefore be a good candidate for refactoring the code
* Checks if you are using THIS.Parent in a form already
* Checks if you are using duplicate methods in column headers which suggest a possible refactoring use for BindEvent.

#### Default Line Rules

* Warns if you are using CTOD on a line of code instead of the DATE() function
* Checks if you have more than 3 .Parent lines on the same line, suggesting possible too many uses of .Parent
