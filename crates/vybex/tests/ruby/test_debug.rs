use super::helpers::compile_ok;

#[test] fn simple_break_if() { compile_ok("break if true\n"); }
#[test] fn while_with_break_if() { compile_ok("while true\n  break if true\nend\n"); }
#[test] fn full_while_break() { compile_ok("x = 0\nwhile true\n  break if x >= 3\n  x += 1\nend\n"); }
#[test] fn while_with_break_only() { compile_ok("while true\n  break\nend\n"); }
#[test] fn while_two_stmts() { compile_ok("while true\n  x = 1\n  break\nend\n"); }
#[test] fn minimal_break_in_while() { compile_ok("while true\nbreak if true\nend\n"); }
#[test] fn two_stmt_in_while() { compile_ok("while true\nx = 1\nbreak if true\nend\n"); }
