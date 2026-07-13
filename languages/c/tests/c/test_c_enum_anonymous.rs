use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn enum_anonymous_basic() { assert_eq!(run_c("enum { A = 1, B = 2 }; int main() { printf(\"%d\", A+B); return 0; }"), vec!["3"]); }
#[test] fn enum_anonymous_implicit_values() { assert_eq!(run_c("enum { A, B, C }; int main() { printf(\"%d\", C); return 0; }"), vec!["2"]); }
#[test] fn enum_anonymous_mixed_values() { assert_eq!(run_c("enum { A = 5, B, C = 10, D }; int main() { printf(\"%d\", B+D); return 0; }"), vec!["17"]); }
#[test] fn enum_anonymous_negative_values() { assert_eq!(run_c("enum { A = -5, B }; int main() { printf(\"%d\", B); return 0; }"), vec!["-4"]); }
#[test] fn enum_anonymous_duplicate_values() { assert_eq!(run_c("enum { A = 1, B = 1 }; int main() { printf(\"%d\", A == B); return 0; }"), vec!["1"]); }
#[test] fn enum_anonymous_typedef() { assert_eq!(run_c("typedef enum { A = 10, B } MyEnum; int main() { MyEnum e = B; printf(\"%d\", e); return 0; }"), vec!["11"]); }
#[test] fn enum_anonymous_in_struct() { assert_eq!(run_c("struct S { enum { A=1, B=2 } e; }; int main() { struct S s; s.e = B; printf(\"%d\", s.e); return 0; }"), vec!["2"]); }
#[test] fn enum_anonymous_in_struct_scope() { assert_eq!(run_c("struct S { enum { A=5 } e; }; int main() { printf(\"%d\", A); return 0; }"), vec!["5"]); } // Enums defined in struct have same scope as struct
#[test] fn enum_anonymous_in_function() { assert_eq!(run_c("int main() { enum { A = 7 }; printf(\"%d\", A); return 0; }"), vec!["7"]); }
#[test] fn enum_anonymous_shadowing() { assert_eq!(run_c("enum { A = 1 }; int main() { enum { A = 2 }; printf(\"%d\", A); return 0; }"), vec!["2"]); }
#[test] fn enum_anonymous_global_local_shadow() { assert_eq!(run_c("int A = 5; int main() { enum { A = 10 }; printf(\"%d\", A); return 0; }"), vec!["10"]); }
#[test] fn enum_anonymous_sizeof() { assert_eq!(run_c("enum { A = 1 }; int main() { printf(\"%d\", (int)sizeof(A)); return 0; }"), vec!["4"]); } // Enums are ints
#[test] fn enum_anonymous_max_int() { assert_eq!(run_c("enum { A = 2147483647 }; int main() { printf(\"%d\", A); return 0; }"), vec!["2147483647"]); }
#[test] fn enum_anonymous_overflow_wraps() { assert_eq!(run_c("enum { A = 2147483647, B }; int main() { printf(\"%d\", B); return 0; }"), vec!["-2147483648"]); } // Implementation defined whether it wraps to negative or uses wider type, commonly wraps to negative or fails. Vybe behavior testing.
#[test] fn enum_anonymous_hex_values() { assert_eq!(run_c("enum { A = 0x10, B = 0x20 }; int main() { printf(\"%d\", A+B); return 0; }"), vec!["48"]); }
