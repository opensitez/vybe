use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn struct_desig_array_basic() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr[1] = 5 }; printf(\"%d\", s.arr[1]); return 0; }"), vec!["5"]); }
#[test] fn struct_desig_array_multiple() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr[0] = 1, .arr[2] = 3 }; printf(\"%d\", s.arr[0] + s.arr[2]); return 0; }"), vec!["4"]); }
#[test] fn struct_desig_array_out_of_order() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr[2] = 3, .arr[0] = 1 }; printf(\"%d\", s.arr[0]); return 0; }"), vec!["1"]); }
#[test] fn struct_desig_array_nested_struct() { assert_eq!(run_c("struct Inner { int a; }; struct Outer { struct Inner arr[2]; }; int main() { struct Outer o = { .arr[1].a = 42 }; printf(\"%d\", o.arr[1].a); return 0; }"), vec!["42"]); }
#[test] fn struct_desig_array_nested_array() { assert_eq!(run_c("struct S { int arr[2][2]; }; int main() { struct S s = { .arr[1][0] = 99 }; printf(\"%d\", s.arr[1][0]); return 0; }"), vec!["99"]); }
#[test] fn struct_desig_array_with_normal_fields() { assert_eq!(run_c("struct S { int a; int arr[2]; int b; }; int main() { struct S s = { .a = 1, .arr[1] = 2, .b = 3 }; printf(\"%d\", s.arr[1] + s.b); return 0; }"), vec!["5"]); }
#[test] fn struct_desig_array_override() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr[0] = 1, .arr[0] = 2 }; printf(\"%d\", s.arr[0]); return 0; }"), vec!["2"]); } // Later designation overrides
#[test] fn struct_desig_array_unspecified_zero() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr[1] = 5 }; printf(\"%d\", s.arr[0]); return 0; }"), vec!["0"]); }
#[test] fn struct_desig_array_gnu_range() { assert_eq!(run_c("struct S { int arr[5]; }; int main() { struct S s = { .arr[1 ... 3] = 7 }; printf(\"%d\", s.arr[2]); return 0; }"), vec!["7"]); } // GNU extension often supported
#[test] fn struct_desig_array_mixed_syntax() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr = {1, 2, 3} }; printf(\"%d\", s.arr[1]); return 0; }"), vec!["2"]); }
#[test] fn struct_desig_array_mixed_syntax_desig_inside() { assert_eq!(run_c("struct S { int arr[3]; }; int main() { struct S s = { .arr = { [1] = 5 } }; printf(\"%d\", s.arr[1]); return 0; }"), vec!["5"]); }
#[test] fn struct_desig_array_pointer_array() { assert_eq!(run_c("struct S { int *arr[2]; }; int main() { int x = 5; struct S s = { .arr[1] = &x }; printf(\"%d\", *s.arr[1]); return 0; }"), vec!["5"]); }
#[test] fn struct_desig_array_struct_array_mixed() { assert_eq!(run_c("struct Inner { int a; }; struct S { struct Inner arr[2]; }; int main() { struct S s = { .arr = { [1] = { .a = 8 } } }; printf(\"%d\", s.arr[1].a); return 0; }"), vec!["8"]); }
#[test] fn struct_desig_array_chars() { assert_eq!(run_c("struct S { char arr[3]; }; int main() { struct S s = { .arr[2] = 'X' }; printf(\"%c\", s.arr[2]); return 0; }"), vec!["X"]); }
#[test] fn struct_desig_array_compound_literal() { assert_eq!(run_c("struct S { int arr[2]; }; int main() { struct S s; s = (struct S){ .arr[1] = 10 }; printf(\"%d\", s.arr[1]); return 0; }"), vec!["10"]); }
