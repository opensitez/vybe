use super::helpers::{run_python, run_python_one, compile_ok};

// Literals & expressions

#[test] fn int_literal() { compile_ok("x = 42\n"); }
#[test] fn float_literal() { compile_ok("x = 3.14\n"); }
#[test] fn string_literal() { compile_ok("x = \"hello\"\n"); }
#[test] fn bool_literal() { compile_ok("x = True\n"); }
#[test] fn none_literal() { compile_ok("x = None\n"); }
#[test] fn list_literal() { compile_ok("x = [1, 2, 3]\n"); }
#[test] fn tuple_literal() { compile_ok("x = (1, 2)\n"); }
#[test] fn dict_literal() { compile_ok("x = {\"a\": 1, \"b\": 2}\n"); }
#[test] fn set_literal() { compile_ok("x = {1, 2, 3}\n"); }
#[test] fn fstring() { compile_ok("name = \"world\"\nx = f\"hello {name}\"\n"); }

// Operators

#[test] fn arithmetic() { compile_ok("x = 1 + 2 * 3 - 4 / 2\n"); }
#[test] fn floor_div_mod() { compile_ok("x = 10 // 3\ny = 10 % 3\n"); }
#[test] fn power() { compile_ok("x = 2 ** 10\n"); }
#[test] fn comparison() { compile_ok("x = 1 < 2\ny = a == b\n"); }
#[test] fn bool_ops() { compile_ok("x = a and b or not c\n"); }
#[test] fn bitwise() { compile_ok("x = a & b | c ^ d\n"); }
#[test] fn unary() { compile_ok("x = -a\ny = +b\nz = ~c\n"); }
#[test] fn ternary() { compile_ok("x = 1 if True else 0\n"); }

// Assignment

#[test] fn augmented_assign() { compile_ok("x = 0\nx += 5\nx -= 1\nx *= 2\nx //= 3\n"); }
#[test] fn tuple_unpacking() { compile_ok("a, b = 1, 2\n"); }

// Runtime tests

#[test]
fn hello_world() {
    let out = run_python("print(\"Hello, World!\")\n");
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn print_number() {
    assert_eq!(run_python_one("print(42)\n"), "42");
}

#[test]
fn print_bool() {
    assert_eq!(run_python_one("print(True)\n"), "true");
}

#[test]
fn var_assignment() {
    let out = run_python("x = 10\ny = 20\nprint(x + y)\n");
    assert_eq!(out, vec!["30"]);
}

#[test]
fn string_concat() {
    assert_eq!(run_python_one("print(\"hello\" + \" \" + \"world\")\n"), "hello world");
}

#[test]
fn arithmetic_runtime() {
    assert_eq!(run_python_one("print(2 + 3 * 4)\n"), "14");
}

#[test]
fn comparison_runtime() {
    assert_eq!(run_python_one("print(5 > 3)\n"), "true");
    assert_eq!(run_python_one("print(5 < 3)\n"), "false");
}

#[test]
fn compound_assignment_runtime() {
    let out = run_python("x = 10\nx += 5\nx -= 3\nx *= 2\nprint(x)\n");
    assert_eq!(out, vec!["24"]);
}

#[test]
fn type_annotation() { compile_ok("x: int = 5\ny: str = 'hello'\n"); }
