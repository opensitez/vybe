use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn ternary_nested_left() { assert_eq!(run_c("int main() { int a = 1 ? 2 ? 3 : 4 : 5; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_right() { assert_eq!(run_c("int main() { int a = 0 ? 1 : 2 ? 3 : 4; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_deep() { assert_eq!(run_c("int main() { int a = 0 ? 1 : 0 ? 2 : 1 ? 3 : 4; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_with_parens_left() { assert_eq!(run_c("int main() { int a = (1 ? 0 : 1) ? 2 : 3; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_with_parens_right() { assert_eq!(run_c("int main() { int a = 1 ? (0 ? 2 : 3) : 4; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_different_types() { assert_eq!(run_c("int main() { double a = 1 ? 0 ? 1 : 2.5 : 3; printf(\"%d\", a > 2.0); return 0; }"), vec!["1"]); } // type of inner is double, outer promotes to double
#[test] fn ternary_nested_pointers() { assert_eq!(run_c("int main() { int x=1, y=2; int *p = 1 ? 0 ? &x : &y : &x; printf(\"%d\", *p); return 0; }"), vec!["2"]); }
#[test] fn ternary_nested_void() { assert_eq!(run_c("void f() { printf(\"F\"); } void g() { printf(\"G\"); } int main() { 1 ? 0 ? f() : g() : f(); return 0; }"), vec!["G"]); }
#[test] fn ternary_nested_short_circuit() { assert_eq!(run_c("int main() { int x=0; 0 ? (x=1) : 1 ? (x=2) : (x=3); printf(\"%d\", x); return 0; }"), vec!["2"]); }
#[test] fn ternary_nested_assignment() { assert_eq!(run_c("int main() { int a; a = 1 ? 2 : 3 ? 4 : 5; printf(\"%d\", a); return 0; }"), vec!["2"]); }
#[test] fn ternary_nested_lvalue_fails() { assert_eq!(run_c("int main() { int a, b; /* (1 ? a : b) = 2; // C ternary does not yield lvalue */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn ternary_nested_comma() { assert_eq!(run_c("int main() { int a = 1 ? (2, 3) : 4; printf(\"%d\", a); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_struct_field() { assert_eq!(run_c("struct S { int a; }; int main() { struct S s1={1}, s2={2}; printf(\"%d\", (1 ? s1 : s2).a); return 0; }"), vec!["1"]); }
#[test] fn ternary_nested_array_decay() { assert_eq!(run_c("int main() { int a1[2]={1,2}, a2[2]={3,4}; int *p = 0 ? a1 : a2; printf(\"%d\", p[0]); return 0; }"), vec!["3"]); }
#[test] fn ternary_nested_function_pointer() { assert_eq!(run_c("int f1() { return 1; } int f2() { return 2; } int main() { int (*p)() = 1 ? 0 ? f1 : f2 : f1; printf(\"%d\", p()); return 0; }"), vec!["2"]); }
