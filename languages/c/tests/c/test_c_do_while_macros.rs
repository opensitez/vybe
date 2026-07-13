use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn do_while_macro_basic() { assert_eq!(run_c("#define WRAP(x) do { printf(\"%d\", x); } while(0)\nint main() { WRAP(1); return 0; }"), vec!["1"]); }
#[test] fn do_while_macro_in_if() { assert_eq!(run_c("#define WRAP(x) do { printf(\"%d\", x); } while(0)\nint main() { if(1) WRAP(2); else WRAP(3); return 0; }"), vec!["2"]); } // Safe in if-else
#[test] fn do_while_macro_multiple_stmts() { assert_eq!(run_c("#define WRAP do { printf(\"A\"); printf(\"B\"); } while(0)\nint main() { WRAP; return 0; }"), vec!["A", "B"]); }
#[test] fn do_while_macro_break() { assert_eq!(run_c("#define WRAP do { printf(\"A\"); break; printf(\"B\"); } while(0)\nint main() { WRAP; return 0; }"), vec!["A"]); } // break escapes do-while
#[test] fn do_while_macro_continue() { assert_eq!(run_c("#define WRAP do { printf(\"A\"); continue; printf(\"B\"); } while(0)\nint main() { WRAP; return 0; }"), vec!["A"]); } // continue goes to condition, which is 0, so exits
#[test] fn do_while_macro_in_for() { assert_eq!(run_c("#define WRAP do { printf(\"A\"); } while(0)\nint main() { for(int i=0; i<2; i++) WRAP; return 0; }"), vec!["A", "A"]); }
#[test] fn do_while_macro_in_while() { assert_eq!(run_c("#define WRAP do { printf(\"A\"); } while(0)\nint main() { int i=0; while(i++ < 2) WRAP; return 0; }"), vec!["A", "A"]); }
#[test] fn do_while_macro_dangling_else_prevention() { assert_eq!(run_c("#define WRAP do { if (0) printf(\"X\"); } while(0)\nint main() { if (1) WRAP; else printf(\"Y\"); return 0; }"), Vec::<&str>::new()); } // Prints nothing, else goes to outer if correctly
#[test] fn do_while_macro_nested() { assert_eq!(run_c("#define INNER do { printf(\"I\"); } while(0)\n#define OUTER do { printf(\"O\"); INNER; } while(0)\nint main() { OUTER; return 0; }"), vec!["O", "I"]); }
#[test] fn do_while_macro_with_declarations() { assert_eq!(run_c("#define SWAP(a, b) do { int tmp = a; a = b; b = tmp; } while(0)\nint main() { int x=1, y=2; SWAP(x, y); printf(\"%d\", x); return 0; }"), vec!["2"]); }
#[test] fn do_while_macro_shadowing() { assert_eq!(run_c("#define M do { int a = 5; printf(\"%d\", a); } while(0)\nint main() { int a = 1; M; printf(\"%d\", a); return 0; }"), vec!["5", "1"]); }
#[test] fn do_while_macro_return() { assert_eq!(run_c("#define M do { return 42; } while(0)\nint main() { M; return 0; }"), vec!["42"]); } // Though return 42 is exit code, our runner might just see it exit
#[test] fn do_while_macro_goto() { assert_eq!(run_c("#define M do { goto L; } while(0)\nint main() { M; printf(\"X\"); L: printf(\"L\"); return 0; }"), vec!["L"]); }
#[test] fn do_while_macro_constant_expression() { assert_eq!(run_c("#define M do { printf(\"C\"); } while(1==0)\nint main() { M; return 0; }"), vec!["C"]); }
#[test] fn do_while_macro_no_semicolon_usage() { assert_eq!(run_c("#define M do { printf(\"M\"); } while(0)\nint main() { M /* no semi here, handled by macro syntax if not needed, but usually is */ ; return 0; }"), vec!["M"]); }
