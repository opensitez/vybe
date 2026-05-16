use super::helpers::{run_ruby, run_ruby_one, compile_ok};

// ── Literals ────────────────────────────────────────────────────────────────

#[test] fn integer()           { compile_ok("x = 42\n"); }
#[test] fn float()             { compile_ok("x = 3.14\n"); }
#[test] fn negative_number()   { compile_ok("x = -5\n"); }
#[test] fn underscore_number() { compile_ok("x = 1_000_000\n"); }
#[test] fn single_quoted()     { compile_ok("x = 'hello'\n"); }
#[test] fn double_quoted()     { compile_ok("x = \"hello\"\n"); }
#[test] fn string_interp()     { compile_ok("name = 'world'\nx = \"hello #{name}\"\n"); }
#[test] fn escape_sequences()  { compile_ok("x = \"line1\\nline2\\ttab\"\n"); }
#[test] fn empty_string()      { compile_ok("x = ''\n"); }
#[test] fn symbol_literal()    { compile_ok("x = :hello\n"); }
#[test] fn symbol_key()        { compile_ok("h = {name: 'Alice'}\n"); }
#[test] fn true_literal()      { compile_ok("x = true\n"); }
#[test] fn false_literal()     { compile_ok("x = false\n"); }
#[test] fn nil_literal()       { compile_ok("x = nil\n"); }
#[test] fn array_literal()     { compile_ok("a = [1, 2, 3]\n"); }
#[test] fn empty_array()       { compile_ok("a = []\n"); }
#[test] fn nested_array()      { compile_ok("a = [[1, 2], [3, 4]]\n"); }
#[test] fn mixed_array()       { compile_ok("a = [1, 'two', true, nil]\n"); }
#[test] fn hash_rocket()       { compile_ok("h = {'a' => 1, 'b' => 2}\n"); }
#[test] fn hash_symbol_keys()  { compile_ok("h = {name: 'Alice', age: 30}\n"); }
#[test] fn empty_hash()        { compile_ok("h = {}\n"); }
#[test] fn inclusive_range()    { compile_ok("r = 1..10\n"); }
#[test] fn exclusive_range()   { compile_ok("r = 1...10\n"); }
#[test] fn local_var()         { compile_ok("x = 5\ny = x + 1\n"); }
#[test] fn global_var()        { compile_ok("$count = 0\n"); }
#[test] fn constant()          { compile_ok("PI = 3.14159\n"); }

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn hello_world() {
    let out = run_ruby("puts 'Hello, World!'\n");
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn print_number() {
    assert_eq!(run_ruby_one("puts 42\n"), "42");
}

#[test]
fn print_bool() {
    assert_eq!(run_ruby_one("puts true\n"), "true");
}

#[test]
fn var_assignment() {
    let out = run_ruby("x = 10\ny = 20\nputs x + y\n");
    assert_eq!(out, vec!["30"]);
}

#[test]
fn string_concat() {
    assert_eq!(run_ruby_one("puts 'hello' + ' ' + 'world'\n"), "hello world");
}

#[test]
fn arithmetic_result() {
    assert_eq!(run_ruby_one("puts 2 + 3 * 4\n"), "14");
}

#[test]
fn nil_value() {
    assert_eq!(run_ruby_one("puts nil\n"), "null");
}
