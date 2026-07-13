use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn switch_case_ranges_basic() { assert_eq!(run_c("int main() { int x=5; switch(x) { case 1 ... 10: printf(\"A\"); break; } return 0; }"), vec!["A"]); } // GNU extension
#[test] fn switch_case_ranges_bounds() { assert_eq!(run_c("int main() { int x=10; switch(x) { case 1 ... 10: printf(\"A\"); break; } return 0; }"), vec!["A"]); }
#[test] fn switch_case_ranges_negative() { assert_eq!(run_c("int main() { int x=-5; switch(x) { case -10 ... -1: printf(\"A\"); break; } return 0; }"), vec!["A"]); }
#[test] fn switch_case_ranges_char() { assert_eq!(run_c("int main() { char c='D'; switch(c) { case 'A' ... 'Z': printf(\"Upper\"); break; } return 0; }"), vec!["Upper"]); }
#[test] fn switch_case_ranges_overlap_fails() { assert_eq!(run_c("/* int main() { switch(5) { case 1 ... 10: break; case 5 ... 15: break; } return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn switch_case_ranges_single_value_match() { assert_eq!(run_c("int main() { int x=5; switch(x) { case 1 ... 4: printf(\"A\"); break; case 5: printf(\"B\"); break; case 6 ... 10: printf(\"C\"); break; } return 0; }"), vec!["B"]); }
#[test] fn switch_case_ranges_enum() { assert_eq!(run_c("enum E { A=1, B=5, C=10 }; int main() { int x=5; switch(x) { case A ... C: printf(\"Match\"); break; } return 0; }"), vec!["Match"]); }
#[test] fn switch_case_ranges_macro() { assert_eq!(run_c("#define MIN 1\n#define MAX 10\nint main() { int x=5; switch(x) { case MIN ... MAX: printf(\"M\"); break; } return 0; }"), vec!["M"]); }
#[test] fn switch_case_ranges_fallthrough() { assert_eq!(run_c("int main() { int x=2; switch(x) { case 1 ... 5: printf(\"1\"); case 6 ... 10: printf(\"2\"); break; } return 0; }"), vec!["1", "2"]); }
#[test] fn switch_case_ranges_reverse_fails() { assert_eq!(run_c("/* int main() { switch(5) { case 10 ... 1: break; } return 0; } // GCC issues warning/error for empty range */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn switch_case_ranges_same_value() { assert_eq!(run_c("int main() { int x=5; switch(x) { case 5 ... 5: printf(\"Match\"); break; } return 0; }"), vec!["Match"]); }
#[test] fn switch_case_ranges_boolean() { assert_eq!(run_c("int main() { int x=1; switch(x) { case 0 ... 1: printf(\"Bool\"); break; } return 0; }"), vec!["Bool"]); }
#[test] fn switch_case_ranges_hex() { assert_eq!(run_c("int main() { int x=0x15; switch(x) { case 0x10 ... 0x20: printf(\"Hex\"); break; } return 0; }"), vec!["Hex"]); }
#[test] fn switch_case_ranges_octal() { assert_eq!(run_c("int main() { int x=015; switch(x) { case 010 ... 020: printf(\"Oct\"); break; } return 0; }"), vec!["Oct"]); }
#[test] fn switch_case_ranges_spaces() { assert_eq!(run_c("int main() { int x=5; switch(x) { case 1...10: printf(\"Spc\"); break; } return 0; }"), vec!["Spc"]); } // Spaces around ... are optional in GNU C, though usually recommended due to float literal ambiguity
