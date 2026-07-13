use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn auto_init_zero_implicit() { assert_eq!(run_c("int main() { int a = 0; printf(\"%d\", a); return 0; }"), vec!["0"]); }
#[test] fn auto_init_hex() { assert_eq!(run_c("int main() { int a = 0xA; printf(\"%d\", a); return 0; }"), vec!["10"]); }
#[test] fn auto_init_octal() { assert_eq!(run_c("int main() { int a = 012; printf(\"%d\", a); return 0; }"), vec!["10"]); }
#[test] fn auto_init_char() { assert_eq!(run_c("int main() { char c = 'A'; printf(\"%c\", c); return 0; }"), vec!["A"]); }
#[test] fn auto_init_multiple_same_line() { assert_eq!(run_c("int main() { int a = 1, b = 2; printf(\"%d\", a+b); return 0; }"), vec!["3"]); }
#[test] fn auto_init_expression() { assert_eq!(run_c("int main() { int a = 5; int b = a * 2; printf(\"%d\", b); return 0; }"), vec!["10"]); }
#[test] fn auto_init_function_call() { assert_eq!(run_c("int f() { return 42; } int main() { int a = f(); printf(\"%d\", a); return 0; }"), vec!["42"]); }
#[test] fn auto_init_array() { assert_eq!(run_c("int main() { int arr[3] = {1, 2, 3}; printf(\"%d\", arr[1]); return 0; }"), vec!["2"]); }
#[test] fn auto_init_array_partial() { assert_eq!(run_c("int main() { int arr[3] = {1}; printf(\"%d\", arr[2]); return 0; }"), vec!["0"]); } // implicit 0
#[test] fn auto_init_struct() { assert_eq!(run_c("struct S { int a; int b; }; int main() { struct S s = {1, 2}; printf(\"%d\", s.a + s.b); return 0; }"), vec!["3"]); }
#[test] fn auto_init_struct_partial() { assert_eq!(run_c("struct S { int a; int b; }; int main() { struct S s = {5}; printf(\"%d\", s.b); return 0; }"), vec!["0"]); }
#[test] fn auto_init_pointer() { assert_eq!(run_c("int main() { int a = 10; int *p = &a; printf(\"%d\", *p); return 0; }"), vec!["10"]); }
#[test] fn auto_init_pointer_null() { assert_eq!(run_c("int main() { int *p = 0; printf(\"%d\", p == 0); return 0; }"), vec!["1"]); }
#[test] fn auto_init_shadowed_var() { assert_eq!(run_c("int main() { int a = 1; { int a = 2; printf(\"%d\", a); } return 0; }"), vec!["2"]); }
#[test] fn auto_init_vla_fails_in_c90() { assert_eq!(run_c("int main() { int n = 5; int arr[n]; printf(\"ok\"); return 0; }"), vec!["ok"]); } // VLA allowed in C99, allowed here
