# VBA Expressions
## ![VBA Expressions](/docs/assets/img/VBA%20Expressions-NewLogo.png)
[![GitHub](https://img.shields.io/github/license/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/blob/master/LICENSE) [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ws-garcia/VBA-Expressions?style=plastic)](https://github.com/ws-garcia/VBA-Expressions/releases/latest) [![Mentioned in Awesome VBA](https://awesome.re/mentioned-badge.svg)](https://github.com/sancarn/awesome-vba)

## Introductory words
Welcome to VBA Expressions, a powerful library designed to extend the capabilities of Visual Basic for Applications (VBA) within Microsoft Office and [LibreOffice environments](https://extensions.libreoffice.org/en/extensions/show/70059). This tool enriches the standard VBA language and LO BASIC with an extensive set of functions for data manipulation, calculation, and analysis, making it ideal for users needing to perform complex operations directly within their spreadsheets, documents, presentatios and much more. 
## ![User Manual](/docs/assets/img/VBAExprManual.pdf)
## Key features
* __Extensive Function Library__: Includes over 100 functions spanning:
    * Mathematical operations
    * Statistical analysis
    * Financial calculations
    * String and date-time manipulations
* __User-Defined Functions (UDFs)__: Create custom functions to meet specific business or analytical needs.
* __Matrix Support__: Handle matrices for advanced mathematical computations or data transformations.
* __Seamless VBA and LO BASIC Integration__: Works directly within VBA and LO BASIC scripts, allowing for automation of complex tasks in Microsoft Excel, Access, Word, and LibreOffice.

## Use cases
### Data Management and Analysis
* __CSV Data Processing__:
    * Filter, join, and transform CSV data directly within VBA and LO BASIC scripts, as seen with the [CSV Interface library](https://github.com/ws-garcia/VBA-CSV-interface)
* __Dynamic Reporting__:
    * Generate dynamic reports by combining data from multiple sources and applying complex filters or calculations.
### Financial Analysis
* __Investment Analysis__: Use financial functions for NPV, IRR, or other investment metrics calculations directly within your code on all supported office ecosystem.
* __Portfolio Management__: Manage and analyze financial data sets with custom calculations for performance metrics.
### Scientific and Engineering Calculations
* __Matrix Operations__: Solve linear systems, perform matrix multiplication for data modeling or simulation.
* __Statistical Analysis__: Conduct statistical tests or data fitting within your Office tools without external software.
### Educational and Technical Documentation
* __Interactive Documents__: Embed calculations or data transformations in educational materials or technical documentation for live demos or explanations.
### Automation of Repetitive Tasks
* __Data Cleaning__: Automate data normalization, cleaning, or validation processes.
* __Batch Processing__: Handle batch operations on data, like updating datasets with new calculations or conditions.

## Advantages
* __Easy to use and integrate__.
* __Basic math operators__: `+` `-` `*` `/` `\` `^` `!`
* __Logical expressions__: `&` (AND), `|` (OR), `||` (XOR)
* __Binary relations__: `<`, `<=`, `<>`, `>=`, `=`, `==`, `>`, `$` (LIKE)
* __Outstanding matrix and statistical functions__: `CHOLESKY`, `MLR` (Multivariate Linear Regression), `FIT` (Curve fitting), `INVERSE`, and a lot more!
* __More than 100 built-in functions__: `Max`, `Sin`, `IRR`, `GAUSS`, `LSQRSOLVE`, `Switch`, `Iff`, `DateDiff`, `Solve`, `fZero`, `Format`...
* __Very flexible and powerful__: variables, constants and user-defined functions (UDFs) support.
* __Implied multiplication for variables, constants and functions__: `5avg(2;abs(-3-7tan(5));9)`, `5(x)` and `x(2)(3)` are valid expressions.
* __Evaluation of arrays of expressions given as text strings, as in Java__: curly brackets must be used to define arrays`{{...};{...}}`
* __Floating point notation input support__: `-5E-5`, `(1.434E3+1000)*2/3.235E-5` are valid inputs.
* __Free of external VBA dependencies__: does not use dll.

## Let's prove it!
## ![Let's prove it!](/docs/assets/img/GeoProblem.png)
The versatility of VBA Expressions allows its users to apply their creativity to solve a wide variety of geometrical problems in anautomated way. To demonstrate this we will solve the problem shown in the figure shown above. We are asked to calculate the area of the largest square inscribed in a right triangle of legs _a_ and _b_. The square corners confguration is: one corner intersecting the leg _a_, another with intersection on leg _b_ and the others corners intersecting the hypotenuse _c_.  First, we define the point _O_ as the square's intersection in leg _b_. The right angle is assumed to be located at the origin of Cartesian plane. In the same way we define _P_ as one intersection of the square with the hypotenuse and _Q_ as the intersection with the leg _a_. Since the figure we are looking for is a square, all sides have a dimension _s_, hence the problem is reduced to ensuring that the magnitudes _OP_ and _OQ_ have the same value. The code to solve the problem is shown below

```vb
Public Function LargestSquareInRATriangle(A As Double, B As Double) As String
    Dim contFunct As String
    Dim evalHelper As VBAexpressions
    Dim xVal As String
    Dim Area As String

    contFunct = "FZERO('DISTANCE(LINESINTERSECT(PERPENDICULAR(" _
                                                    & "{{" & B & ";0};{0;" & A & "}};{{x;0}})" _
                                & ";{{" & B & ";0};{0;" & A & "}})" _
                        & ";{{x;0}})" & "-" _
                & "DISTANCE(LINESINTERSECT(PERPENDICULAR(PERPENDICULAR(" _
                                                        & "{{" & B & ";0};{0;" & A & "}};{{x;0}})" _
                                            & ";{{x;0}})" _
                                & ";{{0;0};{0;" & A & "}})" _
                        & ";{{x;0}})'" _
                    & ";0;" & B & ")"
    Set evalHelper = New VBAexpressions
    With evalHelper
        .Create contFunct
        .Eval
        If .ErrorType = errNone Then
            xVal = Mid(.Result, 2, Len(.Result) - 2)
            .Create "(DISTANCE(LINESINTERSECT(PERPENDICULAR(" _
                                                    & "{{" & B & ";0};{0;" & A & "}};{{x;0}})" _
                                & ";{{" & B & ";0};{0;" & A & "}})" _
                        & ";{{x;0}}))^2", False
            .Eval xVal
            If .ErrorType = errNone Then
                Area = "Area = " & .Result
                LargestSquareInRATriangle = xVal & "; " & Area
            End If
        End If
    End With
End Function
```
Since the magnitude of the segments _OP_ and _OQ_ must be the same, the `FZERO` function can be used to determine the coordinate _O = (x, 0)_ by iteratively computing the value of _x_ that satisfies the continuous equality _OP−OQ = 0_. After execute the call `LargestSquareInRATriangle(5,12)`, the result returned is

```vb
x = 3.14410480351349; Area = 11.6016094276853
```
Another problem that can be solved is shown in the figure below.

![Let's prove it!](/docs/assets/img/ConceptExample.png)

The code solution for the above problem is 

```vb
Public Function IncribedCSCcentersDistance(squareSide As Double) As String
    Dim evalHelper As VBAexpressions
    
    Set evalHelper = New VBAexpressions
    With evalHelper
        .Create "GET('s';" & CStr(squareSide) & ")": .Eval
        .Create "GET('upper.triangle';{{s;0}};{{s;s}};{{0;s}})", False: .Eval
        .Create "GET('lower.triangle';{{0;s}};{{0;-s}};{{s;0}})", False: .Eval
        .Create "GET('ut.inctr';INCENTER(upper.triangle))", False: .Eval
        .Create "GET('lt.inctr';INCENTER(lower.triangle))", False: .Eval
        .Create "GET('inc.dist';ROUND(DISTANCE(ut.inctr;lt.inctr);4))", False: .Eval
        .Create "MROUND(INCIRCLE(upper.triangle);4)", False: .Eval
        Debug.Print "Inscribed upper circle: " & .result
        .Create "MROUND(INCIRCLE(lower.triangle);4)", False: .Eval
        Debug.Print "Inscribed lower circle: " & .result
        IncribedCSCcentersDistance = .VarValue("inc.dist")
    End With
End Function
```
After executing the code, in the console is also printed the inscribed circle center point coordinates and radius magnitudes.

```vb
Inscribed upper circle: {{8.4853;8.4853};{3.5147;3.5147}}
Inscribed lower circle: {{4.9706;0};{4.9706;4.9706}}
```
 
These UDF's are masterpieces demonstrating the power and versatility offered by the latest versions of VBA Expressions!

## Supported expressions

The evaluation approach used is similar to the one we humans use: divide the function into sub-expressions, create a symbolic string to build an expression evaluation flow, split the sub-expressions into chunks of operations (tokens) by tokenization, evaluate all the tokens. 

The module can evaluate mathematical expressions such as:

+ `5*avg(2;abs(-3-7*tan(5));9)-12*pi-e+(7/sin(30)-4!)*min(cos(30);cos(150))`
+ `min(cos(sin(30))+2^2;1)`
+ \* `GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))`

\*`GCD` is an user-defined function (UDF).

Allowed expressions must follow the following grammar:

```
Expression    =     ([{"("}]  SubExpr [{Operator [{"("}] SubExpr [{")"}]}] [{")"}] | {["("] ["{"] List [{";" List}] ["}"] [")"]}) 
SubExpr       =     Token [{Operator Token}]
Token         =     [{Unary}] Argument [(Operator | Function) ["("] [{Unary}] [Argument] [")"]]
Argument      =     (List | Variable | Operand | Literal)
List          =     ["{"] ["{"] SubExpr [{";" SubExpr}] ["}"] ["}"]
Unary         =     "-" | "+" | ~
Literal       =     (Operand | "'"Alphabet"'")
Operand       =     ({Digit} [Decimal] [{Digit}] ["E"("-" | "+"){Digit}] | (True | False))
Variable      =     Alphabet [{Decimal}] [{(Digit | Alphabet)}]
Alphabet      =     "A-Z" | "a-z"
Decimal       =     "." | ","
Digit         =     "0-9"
Operator      =     "+" | "-" | "*" | "/" | "\" | "^" | "%" | "!" | "<" | "<=" | "<>" | ">" | ">=" | "=" | "$" | "&" | "|" | "||"
Function      =     "abs" | "sin" | "cos" | "min" |...|[UDF]
```

## Operators precedence
VBA expressions uses the following precedence rules to evaluate mathematical expressions:

1. `()`                  Grouping: evaluates functions arguments as well.
2. `! - +`               Unary operators: exponentiation is the only operation that violates this. Ex.: `-2 ^ 2 = -4 | (-2) ^ 2 = 4`.
3. `^`                   Exponentiation: Although Excel and Matlab evaluate nested exponentiations from left to right, Google, mathematicians and several modern programming languages, such as Perl, Python and Ruby, evaluate this operation from right to left. VBA expressions also evals in Python way: a^b^c = a^(b^c).
4. `\* / % `             Multiplication, division, modulo: from left to right.
5. `+ -`                 Addition and subtraction: from left to right.
6. `< <= <> >= = == > $` Binary relations.
7. `~`                   Logical negation.
8. `&`                   Logical AND.
9. `||`                  Logical XOR.
10. `|`                  Logical OR.

## Variables
Users can enter variables and set/assign their values for the calculations. Variable names must meet the following requirements:
1. Start with a letter.
2. End in a letter or number. `"x.1"`, `"number1"`, `"value.a"` are valid variable names.
3. A variable named `"A"` is distinct from another variable named `"a"`, since variables are case-sensitive. This rule is broken by the constant `PI`, since `PI=Pi=pi=pI`.
4. The token `"E"` cannot be used as variable due this token is reserved for floating point computation. For example, the expression `"2.5pi+3.5e"` will be evaluated to `~17.3679680`, but a expression like `"2.5pi+3.5E"` will return an error.

## User-defined functions (UDF)
Users can register custom modules to expose and use their functions through the VBAcallBack.cls module. All UDFs must have a single Variant argument that will receive an one-dimensional array of strings (one element for each function argument).

Here is a working example of UDF function creation
```
Sub AddingNewFunctions()
    Dim Evaluator As VBAexpressions
    Dim UDFnames() As Variant
    Dim Result As String
    
    Set Evaluator = New VBAexpressions
    UDFnames() = Array("GCD")
    With Evaluator
        .DeclareUDF UDFnames, "UserDefFunctions"                        'Declare the Greatest Common Divisor function
                                                                        'defined in the UDfunctions class module. This need
                                                                        'an instance in the VBAcallBack class module.
        ' The determinant of a diagonal matrix. It is defined
        ' as the product of the elements in its diagonal.
        ' For our case: 1*2*3=6. (Note that sin(atn(1)*2)=sin(pi/2)=1)
        .Create "GCD(1280;240;100;30*cos(0);10*DET({{sin(atn(1)*2); 0; 0}; {0; 2; 0}; {0; 0; 3}}))"
        Result = .Eval
    End With
End Sub
```
## Working with arrays
VBA expressions can evaluate matrix functions whose arguments are given as arrays/vectors, using a syntax like [Java](https://www.w3schools.com/java/java_arrays.asp). The following expression will calculate, and format to percentage, the internal rate of return (`IRR`) of a cash flow described using a one dimensional array with 5 entries:

`FORMAT(IRR({{-70000;12000;15000;18000;21000}});'Percent')`

However, user-defined array functions need to take care of creating arrays from a string, the `ArrayFromString` method can be used for this purpose.

## Using the code
VBA Expressions is an easy-to-use library, this section shows some examples of how to use the most common properties and methods

$\color{D29922}\textsf{\Large Warning:}$  
>[The library](https://extensions.libreoffice.org/en/extensions/show/70059) only works on LibreOffice version 7.5 or higher and, since there is no 1-1 compatibility between VBA and LO Basic, users must be aware of certain changes required to recover some properties functionality. This applies mainly to those properties related to accessing variables, which were converted into functions to overcome the one-parameter limitation imposed by LO Basic when accessing them, as well as to other properties deprecated due to LO Basic's behaviour in handling class modules. An example of this is the `VarValue` property which was split into two procedures: `GetVarValue` and `LetVarValue`. The rest of the syntax is shared between the two implementations.

```
Sub SimpleMathEval()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "(((((((((((-123.456-654.321)*1.1)*2.2)*3.3)+4.4)+5.5)+6.6)*7.7)*8.8)+9.9)+10.10)"
        If .ReadyToEval Then    'Evaluates only if the expression was successfully parsed.
            .Eval
        End If
    End With
End Sub
Sub LateVariableAssignment()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "Pi.e * 5.2Pie.1 + 3.1Pie"
        If .ReadyToEval Then
            Debug.Print "Variables: "; .CurrentVariables    'Print the list of parsed variables
            .Eval ("Pi.e=1; Pie.1=2; Pie=3")                'Late variable assignment
            Debug.Print .Expression; " = "; .Result; _
                        "; for: "; .CurrentVarValues        'Print stored result, expression and values used in evaluation
        End If
    End With
End Sub
Sub EarlyVariableAssignment()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "Pi.e * 5.2Pie.1 + 3.1Pie"
        If .ReadyToEval Then
            Debug.Print "Variables: "; .CurrentVariables
            .VarValue("Pi.e") = 1
            .ImplicitVarValue("Pie.1") = "2*Pi.e"
            .ImplicitVarValue("Pie") = "Pie.1/3"
            .Eval
            Debug.Print .Expression; " = "; .Result; _
                        "; for: "; .CurrentVarValues
        End If
    End With
End Sub
Sub TrigFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    With Evaluator
        .Create "asin(sin(30))"
        If .ReadyToEval Then
            .Degrees = True               'Eval in degrees
            .Eval
        End If
    End With
End Sub
Sub StringFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create "CONCAT(CHOOSE(1;x;'2nd';'3th';'4th';'5th');'Element';'selected';'/')"
        .Eval ("x='1st'")
    End With
End Sub
Sub LogicalFunctions()
    Dim Evaluator As VBAexpressions
    Set Evaluator = New VBAexpressions
    
    With Evaluator
        .Create "IFF(x > y & x > 0; x; y)"                   
        .Eval("x=70;y=15")                 'This will be evaluated to 70
    End With
End Sub
Sub SolveSystemOfLinearEquations()
    Dim Evaluator As VBAexpressions
    
    Set Evaluator = New VBAexpressions
    With Evaluator
        'Create an array of vectors a, b, c, d and	e
        .Create "SOLVE(ARRAY(a;b;c;d;e);{{'x1';'x2';'x3';'x4';'x5'}};{{100;100;100;100;100}};True)"
		  
        'Define coeficient vectors
        .VarValue("a") = "{4;-1;0;1;0}"
        .VarValue("b") = "{-1;4;-1;0;1}"
        .VarValue("c") = "{0;-1;4;-1;0}"
        .VarValue("d") = "{1;0;-1;4;-1}"
        .VarValue("e") = "{0;1;0;-1;4}"
		  .Eval										'Evaluates to [x1 = 25; x2 = 35.7142857143; x3 = 42.8571428571; x4 = 35.7142857143; x5 = 25]
    End With
End Sub

'@------------------------------------------------------
' Here a list of the new functions and its results
'***********************************ADVANCED MATH FUNCTIONS***************************************************************
''' A={{-2;40};{-1;50};{0;62};{1;58};{2;60}}
''' B={{0;0.1};{0.5;0.45};{1;2.15};{1.5;9.15};{2;40.35};{2.5;180.75}}
''' C={{1;0.01};{2;1};{3;1.15};{4;1.3};{5;1.52};{6;1.84};{7;2.01};{8;2.05};{9;2.3};{10;2.25}}
''' 			 ----------------------------------------------------------------------------
''' 			| Fitting data model with a 4th degree polynomial                            |
''' 			| FIT(A;1;4) : {{62 + 3.6667*x -9.6667*x^2 + 0.3333*x^3 + 1.6667*x^4};{1}}   |
''' 			|                                                                            |
''' 			| Exponential Fitting                                                        |
''' 			| FIT(B;2) : {{0.102*e^(2.9963*x)};{0.9998}}                                 |
''' 			|                                                                            |
''' 			| Logarithmic Fitting                                                        |
''' 			| FIT(C;5) : {{0.9521*ln(x)+0.1049};{0.9752}}                                |
''' 			 ----------------------------------------------------------------------------
'''
''' X={{1;1};{2;2};{3;3};{4;4};{5;1};{6;2};{7;3};{8;4}}
''' Y={{2;4.1;5.8;7.8;5.5;5.2;8.2;11.1}}
''' 			 ------------------------------------------------------------------------------------------
''' 			| Multivariate Linear Regression with regressors/predictors interaction.                   |
''' 			|                                                                                          |
''' 			| MLR(X;Y;True;'X1:X2') : {{0.8542 + 0.4458*X1 + 0.945*X2 + 0.0792*X1*X2};{0.947;0.9072}}  | 
''' 			 ------------------------------------------------------------------------------------------
'''
''' 			 ----------------------------------------------------------------
''' 			| Finding a zero for the given function in the interval -2<=x<=3 |
''' 			|                                                                |
''' 			| FZERO('2x^2+x-12';-2;3) : x = 2.21221445045296                 |
''' 			 ----------------------------------------------------------------
'''
''' a = {1;0;4}; b = {1;1;6}; c = {-3;0;-10}; d = {2;3;4}
''' 			 ------------------------------------------------------------------------------------------------
''' 			| Using LU decomposition for a given matrix                                                      |
''' 			|                                                                                                |
''' 			| LUDECOMP(ARRAY(a;b;c)) :                                                                       |
''' 			|                                                                                                |
''' 			| {{-3;0;-10};{-0.333333333333333;1;2.66666666666667};{-0.333333333333333;0;0.666666666666667}}  |
''' 			 ------------------------------------------------------------------------------------------------
''' 
''' 			 ----------------------------------------------------------------------------------
''' 			| Using LU decomposition for solve linear equations system                         |
''' 			|                                                                                  |
''' 			| LUSOLVE(ARRAY(a;b;c);{{'x';'y';'z'}};{{2;3;4}};True) : x = -18; y = -9; z = 5    |
''' 			 ----------------------------------------------------------------------------------
'''
''' 			 ---------------------------------------------------------------------------
''' 			| Using matrix multiplication for solve a linear equations system           |
''' 			|                                                                           |
''' 			| MMULT(INVERSE(ARRAY(a;b;c));ARRAY(d)) : {{-18};{-9};{5}}                  |     
''' 			| MMULT(ARRAY(a;b;c);INVERSE(ARRAY(a;b;c))) : {{1;0;0};{0;1;0};{0;0;1}}     |
''' 			 ---------------------------------------------------------------------------
'''
''' A={{2;4};{-5;1},{3;-8}};b={{10;-9.5;12}}
''' 			 --------------------------------------------------------------------
''' 		    | Solving overdetermined system of equations using least squares and |
''' 			| the QR decomposition                                               |
''' 			|                                                                    |
''' 			| MROUND(LSQRSOLVE(A;b);4) : {{2.6576;-0.1196}}                      |
''' 			 --------------------------------------------------------------------
'''
'***********************************GEOMETRICAL FUNCTIONS*******************************************************************************
''' DISTANCE({{3.1441;0}};{{4.45415;3.1441}}) = 3.40611153847022
''' LINESINTERSECT({{3.1441;0};{0;1.31004}};{{0;3};{-3;3}}) = {{-4.05590916002565;3}}
''' PARALLEL({{12;0};{0;5}};{{3.1441;0}}) = {{3.1441;0};{0;1.31004166666667}}
''' PERPENDICULAR({{12;0};{0;5}};{{0;0}}) = {{0;0};{1.77514792899408;4.2603550295858}}
''' MROUND(BISECTOR({{-4;-2}};{{-2;6}});4) = {{-3;2};{-2;1.75}}
''' MROUND(INCENTER({{0;0}};{{-4;-2}};{{-2;6}});4) = {{-1.7982;0.7448}}
''' MROUND(INCIRCLE({{0;0}};{{-4;-2}};{{-2;6}});4) = {{-1.7982;0.7448};{1.4704;1.4704}}
''' MROUND(CIRCUMCIRCLE({{0;0}};{{-4;-2}};{{-2;6}});4) = {{-3.5714;2.1429};{4.165;4.165}}
''' MROUND(CIRCLETANG({{-4;3}};SQRT(17);{{2;5}});4) = {{-2.4387;6.8161};{2;5}};{{2;5};{-0.4613;0.8839}}
'***********************************STATISTICAL FUNCTIONS*******************************************************************************
''' ROUND(NORM(0.05);8) = 0.96012239
''' ROUND(CHISQ(4;15);8) = 0.99773734
''' ROUND(GAUSS(0.05);8) = 0.01993881
''' ROUND(ERF(0.05);8) = 0.05637198
''' ROUND(STUDT(0.8;15);8) = 0.43619794
''' ROUND(ANORM(0.75);8) = 0.31863936
''' ROUND(AGAUSS(0.75);8) = 0.67448975
''' ROUND(AERF(0.95);8) = 1.38590382
''' ROUND(ACHISQ(0.75;15);8) = 11.03653766
''' ROUND(FISHF(5.5;1.5;3);8) = 0.21407698
''' ROUND(ASTUDT(0.05;15);8) = 2.13144955
''' ROUND(AFISHF(0.05;1.5;3);8) = 18.55325631
''' ROUND(iBETA(0.5;1;3);8) = 0.875
''' ROUND(BETAINV(0.5;1;3);8) = 0.20629947
'***********************************FINANTIAL FUNCTIONS*********************************************************************************
''' FORMAT(SYD(10000;5000;5;2);'Currency') = '$1,333.33'
''' FORMAT(SLN(10000;0;5);'Currency') = '$2,000.00'
''' FORMAT(RATE(2*12; -250; 5000; 0; 1);'Percent') = '1.66%'
''' FORMAT(PV(0.075/12; 2*12; 250; 0; 0);'Currency') = '($5,555.61)'
''' FORMAT(PPMT(0.06/52; 20; 4*52; 8000; 0; 0);'Currency') = '($34.81)'
''' FORMAT(PMT(0.075/12; 2*12; 5000; 0; 1);'Currency') = '($223.60)'
''' FORMAT(NPV(0.1;{{-10000;3000;4200;6800}});'Currency') = '$1,188.44'
''' FORMAT(NPER(0.0525/1; -200; 1500);'0.00') = '9.78'
''' FORMAT(MIRR({{-7500;3000;5000;1200;4000}};0.05;0.08);'Percent') = '18.74%'
'''
''' Microsoft example. VBA Expressions use a custom solver to compute IRR
''' FORMAT(IRR({{-70000;12000;15000;18000;21000}});'Percent') = '-2.12%'
''' FORMAT(IRR({{-70000;12000;15000;18000;21000;26000}});'Percent') = '8.66%'
''' FORMAT(IRR({{-70000;12000;15000}};true);'Percent') = '-44.35%'					'Find a solution with negative values for IRR
'''
''' FORMAT(IPMT(0.0525/1; 4; 10*1; 6500);'Currency') = '($256.50)'
''' FORMAT(FV(0.0525/1; 10*1; -100; -6500; 0);'Currency') = '$12,115.19'
''' FORMAT(DDB(10000; 5000; 5; 2);'Currency') = '$1,000.00'
'***********************************DATE AND TIME FUNCTIONS*****************************************************************************
''' YEAR(NOW()) = 2022
''' WEEKDAYNAME(1;true;2) = 'lun.'
''' WEEKDAY(NOW()) = 2
''' TIMEVALUE(NOW()) = '7:53:00 a. m.'
''' TIMESERIAL(x;y;z) = '6:45:50 a. m.' for x = 7; y = -15; z = 50
''' MONTH(NOW()) = 10
''' MONTHNAME(x;y) = 'marzo' for x = 3; y = false
''' MONTH(x) = 10 for x = '10/10/2022 7:53:10 a. m.'
''' MINUTE(x) = 53 for x = '10/10/2022 7:53:12 a. m.'
''' HOUR(x) = 7 for x = '10/10/2022 7:53:15 a. m.'
''' DAY(DATE()) = 10
''' DATEVALUE(DATE()) = '10/10/2022'
''' DATESERIAL(2022;x+2;3y) = '21/12/2022' for x = 10; y = 7
''' DATEPART(x;DATE()) = 2022 for x = 'yyyy'
''' DATEPART(x;DATE()) = 10 for x = 'm'
''' DATEPART(x;DATE()) = 10 for x = 'd'
''' DATEPART(x;DATE()) = 4 for x = 'q'
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 3 for x = 'yyyy'; y = 3
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 2 for x = 'q'; y = 2
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 5 for x = 'm'; y = 5
''' DATEDIFF(x;DATE();DATEADD(x;y;DATE())) = 10 for x = 'd'; y = 10
''' DATEADD(x;y;DATE()) = 10/10/2025 for x = 'yyyy'; y = 3
''' DATEADD(x;y;DATE()) = 10/4/2023 for x = 'q'; y = 2
''' DATEADD(x;y;DATE()) = 10/3/2023 for x = 'm'; y = 5
''' DATEADD(x;y;DATE()) = 20/10/2022 for x = 'd'; y = 10
''' DATE() = '10/10/2022'
'***********************************STRING FUNCTIONS******************************************************************************
''' UCASE(x) = ' THIS STRING ' for x = ' This String '
''' TRIM(x) = 'Capi tal' for x = '  Capi tal '
''' RIGHT(2x+20-5+x;2) = '90' for x = 25
''' RIGHT(2x+20-5+x;2) = '09' for x = 98
''' REPLACE(x;'a';'A';1;2) = 'CApitAl' for x = 'Capital'
''' MID(x;4;3) = 'ion' for x = 'Region'
''' LEN(x) = 6 for x = 'Region'
''' LEFT(2x+20-5+x;2) = '90' for x = 25
''' LEFT(2x+20-5+x;2) = '30' for x = 98
''' LCASE(x) = 'this string' for x = 'This String'
''' LCASE(x) = '98' for x = 98
''' FORMAT(x;'Percent') = '98.10%' for x = 0.981
''' CHR(x) = ';' for x = 59
''' ASC(x) = 82 for x = 'Region'
''' SWITCH(x='Asia';1;x='Africa';2;x='Oceania';3) = 1 for x = 'Asia'
'@------------------------------------------------------
```
## Contributing
We welcome contributions to enhance VBA Expressions. Whether it's new functions, bug fixes, or documentation improvements, your input can help make this tool even more powerful:
* Fork the repository.
* Make your changes or additions.
* Submit a pull request.
## Support
For issues, feature requests, or questions, please use the issue tracker on GitHub. We're also open to discussions or contributions to our wiki for more use cases and tutorials.
## Credits
- [x] **Inquisitive knight**: new logo design. Awesome job!

## Tested 
[![Rubberduck](https://user-images.githubusercontent.com/5751684/48656196-a507af80-e9ef-11e8-9c09-1ce3c619c019.png)](https://github.com/rubberduck-vba/Rubberduck/) 

## Licence

Copyright: &copy; 2022-2025  [W. García](https://github.com/ws-garcia/).

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/gpl-3.0.html>.

