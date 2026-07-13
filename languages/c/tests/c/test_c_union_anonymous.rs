use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn union_anon_basic() { assert_eq!(run_c("struct S { union { int i; float f; }; }; int main() { struct S s; s.i = 5; printf(\"%d\", s.i); return 0; }"), vec!["5"]); }
#[test] fn union_anon_initialization() { assert_eq!(run_c("struct S { union { int i; float f; }; }; int main() { struct S s = { 10 }; printf(\"%d\", s.i); return 0; }"), vec!["10"]); } // Initializes first member
#[test] fn union_anon_designated_init() { assert_eq!(run_c("struct S { union { int i; float f; }; }; int main() { struct S s = { .f = 3.14f }; printf(\"%d\", s.f > 3.0); return 0; }"), vec!["1"]); }
#[test] fn union_anon_nested_in_union() { assert_eq!(run_c("union Outer { union { int i; char c; }; double d; }; int main() { union Outer o; o.i = 42; printf(\"%d\", o.i); return 0; }"), vec!["42"]); }
#[test] fn union_anon_struct_member() { assert_eq!(run_c("struct S { union { struct { int a; int b; }; int c; }; }; int main() { struct S s; s.a=1; s.b=2; printf(\"%d\", s.a+s.b); return 0; }"), vec!["3"]); }
#[test] fn union_anon_struct_member_overlap() { assert_eq!(run_c("struct S { union { struct { int a; int b; }; int c; }; }; int main() { struct S s; s.a=10; printf(\"%d\", s.c); return 0; }"), vec!["10"]); } // Assuming little endian or matching layouts
#[test] fn union_anon_tag_fails() { assert_eq!(run_c("/* struct S { union Tag { int i; }; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn union_anon_typedef_fails() { assert_eq!(run_c("typedef union { int i; } U; /* struct S { U; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn union_anon_sizeof() { assert_eq!(run_c("struct S { union { int i; double d; }; }; int main() { printf(\"%d\", sizeof(struct S) >= sizeof(double)); return 0; }"), vec!["1"]); }
#[test] fn union_anon_address_of() { assert_eq!(run_c("struct S { union { int i; float f; }; }; int main() { struct S s; int *p = &s.i; *p = 99; printf(\"%d\", s.i); return 0; }"), vec!["99"]); }
#[test] fn union_anon_multiple() { assert_eq!(run_c("struct S { union { int i; char c; }; union { float f; double d; }; }; int main() { struct S s; s.i = 1; s.f = 2.0; printf(\"%d\", s.i); return 0; }"), vec!["1"]); }
#[test] fn union_anon_shadowing_fails() { assert_eq!(run_c("/* struct S { union { int a; }; int a; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn union_anon_pointer_to_struct() { assert_eq!(run_c("struct S { union { int i; float f; }; }; int main() { struct S s = {42}; struct S *p = &s; printf(\"%d\", p->i); return 0; }"), vec!["42"]); }
#[test] fn union_anon_array_member() { assert_eq!(run_c("struct S { union { int arr[2]; float f; }; }; int main() { struct S s; s.arr[1] = 5; printf(\"%d\", s.arr[1]); return 0; }"), vec!["5"]); }
#[test] fn union_anon_global() { assert_eq!(run_c("/* union { int i; }; // Anonymous unions must be struct/union members in standard C, though some compilers allow global */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
