# Advanced Scripting Framework (ASF) Documentation

## Overview
The Advanced Scripting Framework (ASF) is a lightweight, JavaScript-inspired scripting language embedded in VBA, leveraging the VBA Expressions library (VBAExpr) for computational tasks. This document provides a full detailed grammar in Extended Backus-Naur Form (EBNF) and outlines the key advantages of the implementation.

The grammar covers all core features: variables, control structures, functions (including regular, arrow, anonymous, and multilevel), error handling, output, expressions with operators and literals, and array literals/access with square brackets. Array literals [1,2,3] are converted to VBAExpr matrix notation {1;2;3} for evaluation, and results are reversed (Variant array). Indexing arr[0] is handled as 0-based access on Variant arrays, with assignment arr[1] = 5 supported.

The implementation is a recursive descent parser with an AST-based interpreter, ensuring clarity and extensibility.

## Full Grammar (EBNF)
The grammar is defined in EBNF notation. Terminal symbols are quoted (e.g., "var"), non-terminals are unquoted, `?` denotes optional, `*` zero-or-more, `+` one-or-more,  alternatives, and `[]` grouping.

### Program Structure
```
program ::= statement*

statement ::= varDecl | assignment | ifStmt | whileStmt | forStmt | funcDef | callStmt | returnStmt | breakStmt | continueStmt | printStmt | tryStmt | emptyStmt

emptyStmt ::= ";"
```

- **Example**
```js
var x = 5;
if (x > 0) {
    print('Positive');
}
```

### Declarations and Assignments
```
varDecl ::= "var" identifier "=" expr ";"
assignment ::= lvalue "=" expr ";"

lvalue ::= primary ( "[" expr "]" )*  // Supports chained indexing
```

- **Example**
```js
var y = 10;
y = y + 5;
var arr = [1,2,3];
arr[0] = 100;  // Assignment to array element
```

### Control Flow
```
ifStmt ::= "if" "(" expr ")" block ( "else" ifStmt | "else" block )?

whileStmt ::= "while" "(" expr ")" block

forStmt ::= "for" "(" ( varDecl | ";" )? expr ";" incr ")" block

incr ::= ( identifier "=" )? expr
```

- **Example**
```js
if (x > 5) {
    print('Large');
} else if (x < 5) {
    print('Small');
} else {
    print('Equal');
}

var i = 0;
while (i < 3) {
    print(i);
    i = i + 1;
}

for (var j = 0; j < 3; j = j + 1) {
    print(j);
}
```

### Functions
```
funcDef ::= "function" identifier "(" paramList ")" block

arrowExpr ::= paramList "=>" body ( "=>" arrowExpr )?  // Multilevel support

paramList ::= identifier | "(" ( identifier ( "," identifier )* )? ")"

body ::= block | expr  // Implicit return for expr body

callStmt ::= call ";"

call ::= primary ( "(" argList ")" )*
```

- **Example**
```js
function add(a, b) {
    return a + b;
}
print(add(2, 3));  // Regular function

var double = x => x * 2;  // Arrow function
print(double(4));

var curried = x => y => x + y;  // Multilevel arrow
var add5 = curried(5);
print(add5(3));  // Chained call

print((a, b) => a - b (5, 2));  // Anonymous function call
```

### Error Handling
```
tryStmt ::= "try" block "catch" block
```

- **Example**
```js
try {
    var res = 1 / 0;
    print(res);
} catch {
    print('Division error caught');
}
```

### Output
```
printStmt ::= "print" "(" expr ")" ";"
```

- **Example**
```js
print('Hello ASF');
print(5 + 3);
```

### Statements
```
returnStmt ::= "return" expr? ";"
breakStmt ::= "break" ";"
continueStmt ::= "continue" ";"

block ::= "{" statement* "}"
```

- **Example**
```js
function test() {
    return 42;
}

for (var i = 0; i < 5; i = i + 1) {
    if (i == 3) break;
    print(i);
}

while (true) {
    print('Loop');
    continue;  // Skips rest
    print('Never printed');
}
```

### Expressions
```
expr ::= additive

additive ::= multiplicative ( ( "+" | "-" ) multiplicative )*

multiplicative ::= unary ( ( "*" | "/" | "%" ) unary )*

unary ::= ( "!" )? primary

primary ::= identifier | literal | "(" expr ")" | arrowExpr | arrayLiteral | call

call ::= primary ( "(" argList ")" )*  // Chained calls for anonymous

arrayLiteral ::= "[" ( expr ( "," expr )* )? "]"

argList ::= ( expr ( "," expr )* )?

literal ::= number | string

number ::= [ "-" ] digit+ ( "." digit+ )? [ "e" ( "+" | "-" )? digit+ ]?

string ::= "'" ( ~ "'" | "''" )* "'"

identifier ::= [a-zA-Z] [a-zA-Z0-9_]*  // No leading _
```

- **Example**
```js
var expr = (2 + 3) * 4 / (5 - 1) % 3;
print(expr);  // Arithmetic

var logical = (true && false) || !true;
print(logical);  // Logical (delegated to VBAExpr)

var arrExpr = [1 + 2, 3 * 4];
print(arrExpr[1]);  // Array literal and access
```

### Operators (in expr)
- Arithmetic: `+`, `-`, `*`, `/`, `%`
- Unary: `!`
- Indexing: `[]` (postfix)
- Logical (delegated to VBAExpr): `==`, `!=`, `>`, `<`, `>=`, `<=`, `&&`, `||`

- **Example**
```js
var unary = !true;
print(unary);

var compare = (5 > 3) == true;
print(compare);
```

### Comments (ignored in lexer)
```
comment ::= "//" ~newline* newline | "/*" ( ~ "*/" )* "*/"
```

- **Example**
```js
// Single-line comment
/* Multi-line
comment */
```

## Advantages
- **Portability**: Pure VBA means no installation barriers; runs in any VBA host (Excel, Access, Word). Ideal for enterprise with locked IT—drop classes, run.
- **Expressiveness**: JS-like syntax with arrows/multilevel for functional patterns (currying/recursion), arrays with `[]` for data manipulation, making VBA scripts powerful without boilerplate Subs.
- **Robustness**: AST parser avoids string eval risks; prefixed vars ensure recursion/closures work (tested `factorial=120`, `Fib=5`); `try-catch` handles errors gracefully.
- **Efficiency**: Offloads math/stats to `VBAExpr`; `Variant` arrays fast for access/assign; no parsing for simple expressions.
- **Extensibility**: Class modules allow easy addition (e.g., new `NodeTypes` for switch/classes); shared `exprEval` for custom functions.
- **Security**: Sandboxed execution; no external calls; prefixed locals prevent scope leaks.
- **Simplicity**: `ASFClass.Run(script)` entry point hides complexity; educational for VBA devs learning parsing/AST.

ASF revives VBA as a scripting platform—stunning for 1B+ Office users. In all spirit, it's truth-seeking code: reliable, efficient, and fun.

The core VBA Expressions now Supports logical short-circuiting.