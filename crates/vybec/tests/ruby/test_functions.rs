use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── Basic methods ──────────────────────────────────────────
#[test] fn simple_def() { compile_ok("def greet\n  puts 'hello'\nend"); }
#[test] fn def_with_params() { compile_ok("def add(a, b)\n  a + b\nend"); }
#[test] fn def_with_return() { compile_ok("def square(x)\n  return x * x\nend"); }
#[test] fn def_call() { compile_ok("def greet(name)\n  puts name\nend\ngreet('Alice')"); }

// ── Default parameters ─────────────────────────────────────
#[test] fn default_param() { compile_ok("def greet(name = 'World')\n  puts name\nend"); }
#[test] fn multi_default() { compile_ok("def foo(a, b = 1, c = 2)\n  a + b + c\nend"); }

// ── Splat parameters ───────────────────────────────────────
#[test] fn splat_param() { compile_ok("def foo(*args)\n  puts args\nend"); }
#[test] fn double_splat() { compile_ok("def foo(**opts)\n  puts opts\nend"); }
#[test] fn block_param() { compile_ok("def foo(&block)\n  puts 'has block'\nend"); }

// ── Recursion ──────────────────────────────────────────────
#[test] fn factorial() { compile_ok("def factorial(n)\n  if n <= 1\n    return 1\n  end\n  n * factorial(n - 1)\nend\nfactorial(5)"); }
#[test] fn fibonacci() { compile_ok("def fib(n)\n  return n if n <= 1\n  fib(n - 1) + fib(n - 2)\nend"); }

// ── Lambda / Proc ──────────────────────────────────────────
#[test] fn lambda_arrow() { compile_ok("add = -> (a, b) { a + b }\nadd.call(1, 2)"); }
#[test] fn lambda_keyword() { compile_ok("sq = lambda { |x| x * x }"); }
#[test] fn proc_call() { compile_ok("greet = -> (name) { puts name }\ngreet.call('Alice')"); }

// ── Blocks ─────────────────────────────────────────────────
#[test] fn block_with_each() { compile_ok("[1, 2, 3].each { |x| puts x }"); }
#[test] fn block_do_end() { compile_ok("[1, 2, 3].each do |x|\n  puts x\nend"); }

// ── Yield ──────────────────────────────────────────────────
#[test] fn yield_basic() { compile_ok("def foo\n  yield\nend"); }
#[test] fn yield_with_value() { compile_ok("def foo\n  yield(42)\nend"); }
#[test] fn block_given() { compile_ok("def foo\n  if block_given?\n    yield\n  end\nend"); }

// ── Method chaining ────────────────────────────────────────
#[test] fn chaining() { compile_ok("'hello world'.upcase.split(' ')"); }
