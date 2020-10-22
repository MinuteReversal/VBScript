> ## Visual Basic naming rules
* You must use a letter as the first character.

* You can't use a space, period (.), exclamation mark (!), or the characters @, &, $, # in the name.

* Name can't exceed 255 characters in length.

* Generally, you shouldn't use any names that are the same as the function, statement, method, and intrinsic constant names used in Visual Basic or by the host application. Otherwise you end up shadowing the same keywords in the language. To use an intrinsic language function, statement, or method that conflicts with an assigned name, you must explicitly identify it. Precede the intrinsic function, statement, or method name with the name of the associated type library. For example, if you have a variable called Left, you can only invoke the Left function by using VBA.Left.

* You can't repeat names within the same level of scope. For example, you can't declare two variables named age within the same procedure. However, you can declare a private variable named age and a procedure-level variable named age within the same module.

<div style="background-color:rgb(56,34,93);padding:16px;border-radius:6px;">
Note

Visual Basic isn't case-sensitive, but it preserves the capitalization in the statement where the name is declared.
</div><br/>

> ## docs  
[https://docs.microsoft.com/en-us/office/vba/api/overview/language-reference](https://docs.microsoft.com/en-us/office/vba/api/overview/language-reference)