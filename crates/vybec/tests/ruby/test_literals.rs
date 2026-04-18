use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

fn parse_ok(src: &str) -> bool {
    parse(src).is_ok()
}

// ── Numbers ────────────────────────────────────────────────
#[test] fn integer() { compile_ok("x = 42"); }
#[test] fn float() { compile_ok("x = 3.14"); }
#[test] fn negative_number() { compile_ok("x = -5"); }
#[test] fn underscore_number() { compile_ok("x = 1_000_000"); }

// ── Strings ────────────────────────────────────────────────
#[test] fn single_quoted() { compile_ok("x = 'hello'"); }
#[test] fn double_quoted() { compile_ok("x = \"hello\""); }
#[test] fn string_interpolation() { compile_ok("name = 'world'\nx = \"hello #{name}\""); }
#[test] fn escape_sequences() { compile_ok("x = \"line1\\nline2\\ttab\""); }
#[test] fn empty_string() { compile_ok("x = ''"); }

// ── Symbols ────────────────────────────────────────────────
#[test] fn symbol_literal() { compile_ok("x = :hello"); }
#[test] fn symbol_key() { compile_ok("h = {name: 'Alice'}"); }

// ── Booleans / nil ─────────────────────────────────────────
#[test] fn true_literal() { compile_ok("x = true"); }
#[test] fn false_literal() { compile_ok("x = false"); }
#[test] fn nil_literal() { compile_ok("x = nil"); }

// ── Arrays ─────────────────────────────────────────────────
#[test] fn array_literal() { compile_ok("a = [1, 2, 3]"); }
#[test] fn empty_array() { compile_ok("a = []"); }
#[test] fn nested_array() { compile_ok("a = [[1, 2], [3, 4]]"); }
#[test] fn mixed_array() { compile_ok("a = [1, 'two', true, nil]"); }

// ── Hashes ─────────────────────────────────────────────────
#[test] fn hash_rocket() { compile_ok("h = {'a' => 1, 'b' => 2}"); }
#[test] fn hash_symbol_keys() { compile_ok("h = {name: 'Alice', age: 30}"); }
#[test] fn empty_hash() { compile_ok("h = {}"); }

// ── Ranges ─────────────────────────────────────────────────
#[test] fn inclusive_range() { compile_ok("r = 1..10"); }
#[test] fn exclusive_range() { compile_ok("r = 1...10"); }

// ── Variables ──────────────────────────────────────────────
#[test] fn local_var() { compile_ok("x = 5\ny = x + 1"); }
#[test] fn global_var() { compile_ok("$count = 0"); }
#[test] fn constant() { compile_ok("PI = 3.14159"); }
