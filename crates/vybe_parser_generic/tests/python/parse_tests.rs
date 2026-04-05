/// Python parsing tests — sources from the existing compiler tests.

use vybe_parser_generic::grammar::*;
use vybe_parser_generic::lexer::tokenize;
use vybe_parser_generic::parser::parse;
use vybe_parser_generic::*;

fn must_parse(src: &str) {
    let g = super::grammar::python_grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, true, true);
    if let Err(e) = parse(&tokens, &g) {
        panic!("PARSE FAILED: {}\nSource:\n{}", e, src);
    }
}

fn parse_ok(src: &str) -> Module {
    let g = super::grammar::python_grammar();
    let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, true, true);
    parse(&tokens, &g).unwrap_or_else(|e| panic!("parse failed: {}", e))
}

// ═══════════════════════════════════════════════════════════
// LITERALS (from test_compile_basics.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn int_literal()     { must_parse("x = 42\n"); }
#[test] fn float_literal()   { must_parse("x = 3.14\n"); }
#[test] fn string_literal()  { must_parse("x = \"hello\"\n"); }
#[test] fn bool_literal()    { must_parse("x = True\n"); }
#[test] fn none_literal()    { must_parse("x = None\n"); }
#[test] fn list_literal()    { must_parse("x = [1, 2, 3]\n"); }
#[test] fn dict_literal()    { must_parse("x = {\"a\": 1, \"b\": 2}\n"); }
#[test] fn negative_int()    { must_parse("x = -42\n"); }

// ═══════════════════════════════════════════════════════════
// OPERATORS (from test_compile_basics.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn arithmetic()      { must_parse("x = 1 + 2 * 3 - 4 / 2\n"); }
#[test] fn floor_div_mod()   { must_parse("x = 10 // 3\ny = 10 % 3\n"); }
#[test] fn power()           { must_parse("x = 2 ** 10\n"); }
#[test] fn comparison()      { must_parse("x = 1 < 2\n"); }
#[test] fn boolean_ops()     { must_parse("x = True and False or not True\n"); }
#[test] fn bitwise()         { must_parse("x = 5 & 3\ny = 5 | 3\nz = 5 ^ 3\n"); }
#[test] fn shift()           { must_parse("x = 1 << 3\ny = 8 >> 1\n"); }
#[test] fn unary_neg()       { must_parse("x = -5\ny = +3\nz = ~7\n"); }

// ═══════════════════════════════════════════════════════════
// ASSIGNMENT (from test_compile_basics.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn simple_assign()   { must_parse("x = 42\n"); }
#[test] fn aug_assign_add()  { must_parse("x = 1\nx += 2\n"); }
#[test] fn aug_assign_sub()  { must_parse("x = 10\nx -= 3\n"); }
#[test] fn aug_assign_mul()  { must_parse("x = 5\nx *= 3\n"); }
#[test] fn aug_assign_div()  { must_parse("x = 10\nx /= 4\n"); }

// ═══════════════════════════════════════════════════════════
// CONTROL FLOW (from test_compile_control.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn if_basic() {
    must_parse("if True:\n    x = 1\n");
}
#[test] fn if_else() {
    must_parse("if True:\n    x = 1\nelse:\n    x = 2\n");
}
#[test] fn if_elif() {
    must_parse("x = 2\nif x == 1:\n    y = 'one'\nelif x == 2:\n    y = 'two'\nelse:\n    y = 'other'\n");
}
#[test] fn while_basic() {
    must_parse("i = 0\nwhile i < 5:\n    i = i + 1\n");
}
#[test] fn for_range() {
    must_parse("for i in range(5):\n    x = i\n");
}
#[test] fn for_list() {
    must_parse("for x in [1, 2, 3]:\n    y = x\n");
}
#[test] fn break_stmt() {
    must_parse("while True:\n    break\n");
}
#[test] fn continue_stmt() {
    must_parse("for i in range(10):\n    continue\n");
}
#[test] fn pass_stmt() {
    must_parse("if True:\n    pass\n");
}

// ═══════════════════════════════════════════════════════════
// FUNCTIONS (from test_compile_functions.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn def_basic() {
    must_parse("def add(a, b):\n    return a + b\n");
}
#[test] fn def_no_params() {
    must_parse("def greet():\n    return 'hello'\n");
}
#[test] fn def_default_param() {
    must_parse("def greet(name='world'):\n    return 'hello ' + name\n");
}
#[test] fn def_nested() {
    must_parse("def outer(x):\n    def inner(y):\n        return y * 2\n    return inner(x)\n");
}
#[test] fn def_recursive() {
    must_parse("def fact(n):\n    if n <= 1:\n        return 1\n    return n * fact(n - 1)\n");
}
#[test] fn lambda_basic() {
    must_parse("f = lambda x: x + 1\n");
}

// ═══════════════════════════════════════════════════════════
// CLASSES (from test_compile_classes.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn class_basic() {
    must_parse("class Animal:\n    def __init__(self, name):\n        self.name = name\n");
}
#[test] fn class_inheritance() {
    must_parse("class Animal:\n    pass\nclass Dog(Animal):\n    pass\n");
}
#[test] fn class_method() {
    must_parse("class Foo:\n    def bar(self):\n        return 42\n");
}

// ═══════════════════════════════════════════════════════════
// EXCEPTIONS (from test_compile_exceptions.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn try_except() {
    must_parse("try:\n    x = 1\nexcept:\n    x = 0\n");
}
#[test] fn try_except_finally() {
    must_parse("try:\n    x = 1\nexcept:\n    x = 0\nfinally:\n    y = 2\n");
}
#[test] fn raise_basic() {
    must_parse("raise Exception('error')\n");
}

// ═══════════════════════════════════════════════════════════
// BUILTINS (from test_compile_builtins.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn call_print()      { must_parse("print('hello')\n"); }
#[test] fn call_len()        { must_parse("x = len([1, 2, 3])\n"); }
#[test] fn call_range()      { must_parse("x = range(10)\n"); }
#[test] fn call_str()        { must_parse("x = str(42)\n"); }
#[test] fn call_int()        { must_parse("x = int('42')\n"); }
#[test] fn call_nested()     { must_parse("print(len('hello'))\n"); }
#[test] fn method_call()     { must_parse("x = 'hello'.upper()\n"); }
#[test] fn index_access()    { must_parse("x = [1, 2, 3]\ny = x[0]\n"); }

// ═══════════════════════════════════════════════════════════
// PROGRAMS (from test_compile_programs.rs)
// ═══════════════════════════════════════════════════════════

#[test] fn prog_fibonacci() {
    must_parse("def fib(n):\n    if n <= 1:\n        return n\n    return fib(n - 1) + fib(n - 2)\nprint(fib(10))\n");
}

#[test] fn prog_fizzbuzz() {
    must_parse("for i in range(1, 16):\n    if i % 15 == 0:\n        print('FizzBuzz')\n    elif i % 3 == 0:\n        print('Fizz')\n    elif i % 5 == 0:\n        print('Buzz')\n    else:\n        print(i)\n");
}

// ═══════════════════════════════════════════════════════════
// IMPORTS
// ═══════════════════════════════════════════════════════════

#[test] fn import_basic()    { must_parse("import math\n"); }
#[test] fn from_import()     { must_parse("from os import path\n"); }

// ═══════════════════════════════════════════════════════════
// DECORATORS
// ═══════════════════════════════════════════════════════════

#[test] fn decorator_basic() {
    must_parse("@staticmethod\ndef foo():\n    pass\n");
}

// ═══════════════════════════════════════════════════════════
// AST STRUCTURE CHECKS
// ═══════════════════════════════════════════════════════════

#[test] fn ast_function_def() {
    let m = parse_ok("def add(a, b):\n    return a + b\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::FunctionDecl { .. })));
}

#[test] fn ast_class_def() {
    let m = parse_ok("class Foo:\n    pass\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::ClassDecl { .. })));
}

#[test] fn ast_if_stmt() {
    let m = parse_ok("if True:\n    x = 1\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::If { .. })));
}

#[test] fn ast_for_in() {
    let m = parse_ok("for x in [1, 2]:\n    pass\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::ForIn { .. })));
}

#[test] fn ast_while() {
    let m = parse_ok("while True:\n    break\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::While { .. })));
}

#[test] fn ast_assign() {
    let m = parse_ok("x = 42\n");
    assert!(m.body.iter().any(|s| matches!(s.kind, StmtKind::Assign { .. })));
}
