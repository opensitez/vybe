use super::helpers::{compile_ok, run_ruby, run_ruby_one};

// ── Method definitions (compile) ────────────────────────────────────────────

#[test]
fn def_no_params() {
    compile_ok("def greet\n  puts 'hi'\nend\n");
}
#[test]
fn def_params() {
    compile_ok("def add(a, b)\n  a + b\nend\n");
}
#[test]
fn def_defaults() {
    compile_ok("def greet(name = 'world')\n  puts name\nend\n");
}
#[test]
fn def_splat() {
    compile_ok("def show(*args)\n  puts args\nend\n");
}
#[test]
fn def_return() {
    compile_ok("def five\n  return 5\nend\n");
}
#[test]
fn def_implicit_ret() {
    compile_ok("def five\n  5\nend\n");
}
#[test]
fn def_multi_params() {
    compile_ok("def calc(a, b, c = 0)\n  a + b + c\nend\n");
}
#[test]
fn def_with_block() {
    compile_ok("def each_item(&block)\n  block.call(1)\nend\n");
}

// ── Lambda / Proc ───────────────────────────────────────────────────────────

#[test]
fn lambda_arrow() {
    compile_ok("f = ->(x) { x * 2 }\n");
}
#[test]
fn lambda_call() {
    compile_ok("f = ->(x) { x * 2 }\nputs f.call(5)\n");
}
#[test]
fn proc_new() {
    compile_ok("p = Proc.new { |x| x + 1 }\n");
}

// ── Blocks ──────────────────────────────────────────────────────────────────

#[test]
fn block_do_end() {
    compile_ok("[1, 2, 3].each do |x|\n  puts x\nend\n");
}
#[test]
fn block_braces() {
    compile_ok("[1, 2, 3].each { |x| puts x }\n");
}

// ── Yield ───────────────────────────────────────────────────────────────────

#[test]
fn yield_simple() {
    compile_ok("def foo\n  yield\nend\n");
}
#[test]
fn yield_with_val() {
    compile_ok("def foo\n  yield 42\nend\n");
}

// ── Recursion ───────────────────────────────────────────────────────────────

#[test]
fn recursion() {
    compile_ok("def fact(n)\n  if n <= 1\n    return 1\n  end\n  n * fact(n - 1)\nend\n");
}

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn def_call_runtime() {
    let out = run_ruby("def greet\n  puts 'hello'\nend\ngreet\n");
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn def_params_runtime() {
    let out = run_ruby("def add(a, b)\n  puts a + b\nend\nadd(3, 4)\n");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn def_default_runtime() {
    let out =
        run_ruby("def greet(name = 'world')\n  puts 'hello ' + name\nend\ngreet\ngreet('Ruby')\n");
    assert_eq!(out, vec!["hello world", "hello Ruby"]);
}

#[test]
fn def_return_runtime() {
    let out = run_ruby("def five\n  return 5\nend\nputs five()\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn def_implicit_return_runtime() {
    let out = run_ruby("def five\n  5\nend\nputs five()\n");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn recursion_runtime() {
    let out = run_ruby(
        "def fact(n)\n  if n <= 1\n    return 1\n  end\n  n * fact(n - 1)\nend\nputs fact(5)\n",
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn lambda_runtime() {
    let out = run_ruby("f = ->(x) { x * 2 }\nputs f.call(5)\n");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn block_each_runtime() {
    let out = run_ruby("[1, 2, 3].each do |x|\n  puts x\nend\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}
