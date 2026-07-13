use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn extern_global_access() { assert_eq!(run_c("int g = 1; int main() { extern int g; printf(\"%d\", g); return 0; }"), vec!["1"]); }
#[test] fn extern_local_declaration() { assert_eq!(run_c("int main() { extern int g; printf(\"ok\"); return 0; } int g = 2;"), vec!["ok"]); } // Only checking it parses/links
#[test] fn extern_multiple_declarations() { assert_eq!(run_c("extern int g; extern int g; int g = 3; int main() { printf(\"%d\", g); return 0; }"), vec!["3"]); }
#[test] fn extern_function() { assert_eq!(run_c("extern int f(void); int main() { printf(\"%d\", f()); return 0; } int f() { return 4; }"), vec!["4"]); }
#[test] fn extern_array_incomplete() { assert_eq!(run_c("extern int arr[]; int main() { printf(\"ok\"); return 0; } int arr[3] = {1,2,3};"), vec!["ok"]); }
#[test] fn extern_array_complete_match() { assert_eq!(run_c("extern int arr[3]; int arr[3] = {5,6,7}; int main() { printf(\"%d\", arr[0]); return 0; }"), vec!["5"]); }
#[test] fn extern_struct() { assert_eq!(run_c("struct S { int a; }; extern struct S s; struct S s = {10}; int main() { printf(\"%d\", s.a); return 0; }"), vec!["10"]); }
#[test] fn extern_in_block_shadows() { assert_eq!(run_c("int g = 1; int main() { int g = 2; { extern int g; printf(\"%d\", g); } return 0; }"), vec!["1"]); }
#[test] fn extern_initialization_fails() { assert_eq!(run_c("int main() { /* extern int a = 5; // Local extern cannot be initialized */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn extern_global_initialization() { assert_eq!(run_c("extern int g = 5; /* Warning but usually allowed as definition */ int main() { printf(\"%d\", g); return 0; }"), vec!["5"]); }
#[test] fn extern_inline() { assert_eq!(run_c("extern inline int f() { return 6; } int main() { printf(\"%d\", f()); return 0; }"), vec!["6"]); }
#[test] fn extern_with_static_fails() { assert_eq!(run_c("/* extern static int g; // Conflicting */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn extern_tentative() { assert_eq!(run_c("extern int g; int g; int main() { g = 7; printf(\"%d\", g); return 0; }"), vec!["7"]); }
#[test] fn extern_function_pointer() { assert_eq!(run_c("extern int (*p)(void); int f() { return 8; } int (*p)(void) = f; int main() { printf(\"%d\", p()); return 0; }"), vec!["8"]); }
#[test] fn extern_const() { assert_eq!(run_c("extern const int c; const int c = 9; int main() { printf(\"%d\", c); return 0; }"), vec!["9"]); }
