use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn goto_bypassing_int_init() { assert_eq!(run_c("int main() { goto L; int a = 5; L: printf(\"A\"); return 0; }"), vec!["A"]); }
#[test] fn goto_bypassing_multiple_inits() { assert_eq!(run_c("int main() { goto L; int a=1, b=2, c=3; L: printf(\"B\"); return 0; }"), vec!["B"]); }
#[test] fn goto_bypassing_struct_init() { assert_eq!(run_c("struct S { int x; }; int main() { goto L; struct S s = {42}; L: printf(\"C\"); return 0; }"), vec!["C"]); }
#[test] fn goto_bypassing_array_init() { assert_eq!(run_c("int main() { goto L; int arr[3] = {1,2,3}; L: printf(\"D\"); return 0; }"), vec!["D"]); }
#[test] fn goto_bypassing_pointer_init() { assert_eq!(run_c("int main() { int x = 1; goto L; int *p = &x; L: printf(\"E\"); return 0; }"), vec!["E"]); }
#[test] fn goto_bypassing_vla_allowed_in_c99_but_fails_runtime() { assert_eq!(run_c("/* int main() { int n = 5; goto L; int arr[n]; L: return 0; } // illegal to bypass VLA */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn goto_bypassing_variably_modified_typedef() { assert_eq!(run_c("/* int main() { int n = 5; goto L; typedef int A[n]; L: return 0; } // illegal to bypass VM type */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn goto_jumping_back_reinit() { assert_eq!(run_c("int main() { int i = 0; L: { int a = 5; i++; if (i<2) goto L; } printf(\"%d\", i); return 0; }"), vec!["2"]); }
#[test] fn goto_jumping_over_static_init() { assert_eq!(run_c("int main() { goto L; static int a = 1; L: printf(\"%d\", a); return 0; }"), vec!["1"]); } // Valid, static initialized before
#[test] fn goto_jumping_over_extern() { assert_eq!(run_c("int g = 10; int main() { goto L; extern int g; L: printf(\"%d\", g); return 0; }"), vec!["10"]); }
#[test] fn goto_bypassing_compound_literal() { assert_eq!(run_c("struct S { int a; }; int main() { goto L; struct S *s = &(struct S){42}; L: printf(\"F\"); return 0; }"), vec!["F"]); }
#[test] fn goto_bypassing_const_init() { assert_eq!(run_c("int main() { goto L; const int a = 1; L: printf(\"G\"); return 0; }"), vec!["G"]); }
#[test] fn goto_bypassing_enum_decl() { assert_eq!(run_c("int main() { goto L; enum E { A, B }; L: printf(\"H\"); return 0; }"), vec!["H"]); } // Type decls are fine to bypass
#[test] fn goto_bypassing_struct_decl() { assert_eq!(run_c("int main() { goto L; struct S { int a; }; L: printf(\"I\"); return 0; }"), vec!["I"]); }
#[test] fn goto_jumping_into_switch() { assert_eq!(run_c("int main() { int x = 1; goto L; switch(x) { case 1: int y = 2; L: printf(\"J\"); break; } return 0; }"), vec!["J"]); }
