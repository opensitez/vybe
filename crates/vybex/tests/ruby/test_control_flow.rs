use super::helpers::{run_ruby, run_ruby_one, compile_ok};

// ── If / Elsif / Else (compile) ─────────────────────────────────────────────

#[test] fn if_simple()   { compile_ok("if true\n  puts 'yes'\nend\n"); }
#[test] fn if_else()     { compile_ok("if false\n  puts 'no'\nelse\n  puts 'yes'\nend\n"); }
#[test] fn if_elsif()    { compile_ok("x = 2\nif x == 1\n  puts 'a'\nelsif x == 2\n  puts 'b'\nelse\n  puts 'c'\nend\n"); }
#[test] fn unless_simple() { compile_ok("unless false\n  puts 'yes'\nend\n"); }
#[test] fn unless_else() { compile_ok("unless true\n  puts 'no'\nelse\n  puts 'yes'\nend\n"); }

// ── Modifier forms ──────────────────────────────────────────────────────────

#[test] fn modifier_if()     { compile_ok("puts 'yes' if true\n"); }
#[test] fn modifier_unless() { compile_ok("puts 'yes' unless false\n"); }
#[test] fn modifier_while()  { compile_ok("x = 0\nx += 1 while x < 5\n"); }
#[test] fn modifier_until()  { compile_ok("x = 0\nx += 1 until x == 5\n"); }

// ── While / Until ───────────────────────────────────────────────────────────

#[test] fn while_loop()  { compile_ok("x = 0\nwhile x < 5\n  x += 1\nend\n"); }
#[test] fn until_loop()  { compile_ok("x = 0\nuntil x == 5\n  x += 1\nend\n"); }
#[test] fn loop_stmt()   { compile_ok("i = 0\nloop do\n  break if i >= 3\n  i += 1\nend\n"); }

// ── For ─────────────────────────────────────────────────────────────────────

#[test] fn for_range()   { compile_ok("for i in 1..5\n  puts i\nend\n"); }
#[test] fn for_array()   { compile_ok("for x in [1, 2, 3]\n  puts x\nend\n"); }

// ── Case / When ─────────────────────────────────────────────────────────────

#[test] fn case_when()   { compile_ok("x = 2\ncase x\nwhen 1\n  puts 'one'\nwhen 2\n  puts 'two'\nelse\n  puts 'other'\nend\n"); }

// ── Begin / Rescue / Ensure ─────────────────────────────────────────────────

#[test] fn begin_rescue() {
    compile_ok("begin\n  x = 1 / 0\nrescue\n  puts 'error'\nend\n");
}

#[test] fn begin_rescue_ensure() {
    compile_ok("begin\n  x = 1\nrescue => e\n  puts e\nensure\n  puts 'done'\nend\n");
}

// ── Break / Next ────────────────────────────────────────────────────────────

#[test] fn break_in_while() { compile_ok("x = 0\nwhile true\n  break if x >= 3\n  x += 1\nend\n"); }
#[test] fn next_in_while()  { compile_ok("x = 0\nwhile x < 10\n  x += 1\n  next if x % 2 == 0\n  puts x\nend\n"); }

// ── Runtime ─────────────────────────────────────────────────────────────────

#[test]
fn if_runtime() {
    assert_eq!(run_ruby_one("if true\n  puts 'yes'\nend\n"), "yes");
}

#[test]
fn if_else_runtime() {
    assert_eq!(run_ruby_one("if false\n  puts 'no'\nelse\n  puts 'yes'\nend\n"), "yes");
}

#[test]
fn if_elsif_runtime() {
    let out = run_ruby("x = 2\nif x == 1\n  puts 'a'\nelsif x == 2\n  puts 'b'\nelse\n  puts 'c'\nend\n");
    assert_eq!(out, vec!["b"]);
}

#[test]
fn unless_runtime() {
    assert_eq!(run_ruby_one("unless false\n  puts 'yes'\nend\n"), "yes");
}

#[test]
fn while_runtime() {
    let out = run_ruby("x = 0\nwhile x < 3\n  puts x\n  x += 1\nend\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn until_runtime() {
    let out = run_ruby("x = 0\nuntil x == 3\n  puts x\n  x += 1\nend\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_range_runtime() {
    let out = run_ruby("for i in 1..3\n  puts i\nend\n");
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn case_when_runtime() {
    let out = run_ruby("x = 2\ncase x\nwhen 1\n  puts 'one'\nwhen 2\n  puts 'two'\nelse\n  puts 'other'\nend\n");
    assert_eq!(out, vec!["two"]);
}

#[test]
fn modifier_if_runtime() {
    assert_eq!(run_ruby_one("puts 'yes' if true\n"), "yes");
}

#[test]
fn modifier_unless_runtime() {
    assert_eq!(run_ruby_one("puts 'yes' unless false\n"), "yes");
}

#[test]
fn break_runtime() {
    let out = run_ruby("x = 0\nwhile true\n  break if x >= 3\n  puts x\n  x += 1\nend\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn next_runtime() {
    let out = run_ruby("for i in 1..5\n  next if i % 2 == 0\n  puts i\nend\n");
    assert_eq!(out, vec!["1", "3", "5"]);
}

#[test]
fn begin_rescue_runtime() {
    let out = run_ruby("begin\n  raise 'oops'\nrescue => e\n  puts 'caught'\nend\n");
    assert_eq!(out, vec!["caught"]);
}
