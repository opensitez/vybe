use vybec::parser_python::parse;
use vybec::compiler_python::Compiler;

fn compile_ok(src: &str) {
    let module = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&module);
    assert!(res.is_ok(), "compile failed: {:?}", res.err());
    assert!(!res.unwrap().is_empty());
}

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
