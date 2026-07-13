use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn ternary_gnu_omitted_basic() { assert_eq!(run_c("int main() { int a = 5 ?: 10; printf(\"%d\", a); return 0; }"), vec!["5"]); }
#[test] fn ternary_gnu_omitted_false() { assert_eq!(run_c("int main() { int a = 0 ?: 10; printf(\"%d\", a); return 0; }"), vec!["10"]); }
#[test] fn ternary_gnu_omitted_side_effects() { assert_eq!(run_c("int f(int *x) { (*x)++; return *x; } int main() { int x = 0; int a = f(&x) ?: 10; printf(\"%d\", a); return 0; }"), vec!["1"]); } // Evaluated only once
#[test] fn ternary_gnu_omitted_pointers() { assert_eq!(run_c("int main() { int x=42; int *p = &x ?: 0; printf(\"%d\", *p); return 0; }"), vec!["42"]); }
#[test] fn ternary_gnu_omitted_null_pointer() { assert_eq!(run_c("int main() { int x=99; int *p = 0 ?: &x; printf(\"%d\", *p); return 0; }"), vec!["99"]); }
#[test] fn ternary_gnu_omitted_struct_fails() { assert_eq!(run_c("struct S { int a; }; int main() { struct S s={1}; /* struct S res = s ?: s; // Condition must be scalar */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn ternary_gnu_omitted_float() { assert_eq!(run_c("int main() { double d = 0.0 ?: 3.14; printf(\"%d\", d > 3.0); return 0; }"), vec!["1"]); }
#[test] fn ternary_gnu_omitted_function_pointer() { assert_eq!(run_c("int f() { return 1; } int main() { int (*p)() = f ?: 0; printf(\"%d\", p()); return 0; }"), vec!["1"]); }
#[test] fn ternary_gnu_omitted_assignment() { assert_eq!(run_c("int main() { int a=0, b=5; a = a ?: b; printf(\"%d\", a); return 0; }"), vec!["5"]); }
#[test] fn ternary_gnu_omitted_chain() { assert_eq!(run_c("int main() { int a = 0 ?: 0 ?: 7; printf(\"%d\", a); return 0; }"), vec!["7"]); }
#[test] fn ternary_gnu_omitted_promotion() { assert_eq!(run_c("int main() { char c = 'A'; int i = c ?: 1000; printf(\"%d\", i); return 0; }"), vec!["65"]); } // Promotes correctly
#[test] fn ternary_gnu_omitted_void_fails() { assert_eq!(run_c("void f() {} int main() { /* f() ?: 1; // void cannot be tested */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn ternary_gnu_omitted_macro() { assert_eq!(run_c("#define DEFAULT(x, def) ((x) ?: (def))\nint main() { printf(\"%d\", DEFAULT(0, 88)); return 0; }"), vec!["88"]); }
#[test] fn ternary_gnu_omitted_logical() { assert_eq!(run_c("int main() { int a = (1 && 0) ?: 2; printf(\"%d\", a); return 0; }"), vec!["2"]); }
#[test] fn ternary_gnu_omitted_with_comma() { assert_eq!(run_c("int main() { int a = (0, 3) ?: 4; printf(\"%d\", a); return 0; }"), vec!["3"]); }
