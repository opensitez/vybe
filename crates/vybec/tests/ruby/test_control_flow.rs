use vybec::parser_ruby::parse;
use vybec::compiler_ruby::Compiler;

fn compile_ok(src: &str) {
    let program = parse(src).expect("parse failed");
    let mut c = Compiler::new();
    let res = c.compile(&program);
    assert!(res.is_ok(), "compile failed for:\n{}\nerror: {:?}", src, res.err());
}

// ── If/elsif/else ──────────────────────────────────────────
#[test] fn if_basic() { compile_ok("if true\n  puts 'yes'\nend"); }
#[test] fn if_else() { compile_ok("if false\n  puts 'no'\nelse\n  puts 'yes'\nend"); }
#[test] fn if_elsif() { compile_ok("x = 5\nif x > 10\n  puts 'big'\nelsif x > 3\n  puts 'medium'\nelse\n  puts 'small'\nend"); }
#[test] fn if_then() { compile_ok("if true then puts 'yes' end"); }

// ── Unless ─────────────────────────────────────────────────
#[test] fn unless_basic() { compile_ok("unless false\n  puts 'yes'\nend"); }
#[test] fn unless_else() { compile_ok("unless true\n  puts 'no'\nelse\n  puts 'yes'\nend"); }

// ── Modifier if/unless ─────────────────────────────────────
#[test] fn modifier_if() { compile_ok("puts 'hello' if true"); }
#[test] fn modifier_unless() { compile_ok("puts 'hello' unless false"); }

// ── While ──────────────────────────────────────────────────
#[test] fn while_loop() { compile_ok("i = 0\nwhile i < 10\n  i += 1\nend"); }
#[test] fn modifier_while() { compile_ok("i = 0\ni += 1 while i < 10"); }

// ── Until ──────────────────────────────────────────────────
#[test] fn until_loop() { compile_ok("i = 0\nuntil i >= 10\n  i += 1\nend"); }
#[test] fn modifier_until() { compile_ok("i = 0\ni += 1 until i >= 10"); }

// ── For/in ─────────────────────────────────────────────────
#[test] fn for_in() { compile_ok("for x in [1, 2, 3]\n  puts x\nend"); }

// ── Case/when ──────────────────────────────────────────────
#[test] fn case_when() { compile_ok("x = 5\ncase x\nwhen 1\n  puts 'one'\nwhen 5\n  puts 'five'\nelse\n  puts 'other'\nend"); }
#[test] fn case_multi_when() { compile_ok("x = 3\ncase x\nwhen 1, 2, 3\n  puts 'small'\nwhen 4, 5, 6\n  puts 'medium'\nend"); }

// ── Break / Next ───────────────────────────────────────────
#[test] fn break_in_while() { compile_ok("i = 0\nwhile true\n  break if i > 5\n  i += 1\nend"); }
#[test] fn next_in_while() { compile_ok("i = 0\nwhile i < 10\n  i += 1\n  next if i == 5\n  puts i\nend"); }

// ── Return ─────────────────────────────────────────────────
#[test] fn return_value() { compile_ok("def foo\n  return 42\nend"); }
#[test] fn return_nil() { compile_ok("def foo\n  return\nend"); }

// ── Begin/rescue/ensure ────────────────────────────────────
#[test] fn begin_rescue() { compile_ok("begin\n  raise 'oops'\nrescue => e\n  puts e\nend"); }
#[test] fn begin_rescue_ensure() { compile_ok("begin\n  x = 1\nrescue\n  puts 'error'\nensure\n  puts 'done'\nend"); }
#[test] fn rescue_type() { compile_ok("begin\n  raise 'oops'\nrescue RuntimeError => e\n  puts e\nend"); }
