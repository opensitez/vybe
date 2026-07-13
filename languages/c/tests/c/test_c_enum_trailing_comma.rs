use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn enum_trailing_comma_basic() { assert_eq!(run_c("enum E { A, B, }; int main() { printf(\"%d\", B); return 0; }"), vec!["1"]); } // C99 allows trailing comma
#[test] fn enum_trailing_comma_with_values() { assert_eq!(run_c("enum E { A=5, B=6, }; int main() { printf(\"%d\", B); return 0; }"), vec!["6"]); }
#[test] fn enum_trailing_comma_anonymous() { assert_eq!(run_c("enum { A=1, }; int main() { printf(\"%d\", A); return 0; }"), vec!["1"]); }
#[test] fn enum_trailing_comma_single_element() { assert_eq!(run_c("enum E { A, }; int main() { printf(\"%d\", A); return 0; }"), vec!["0"]); }
#[test] fn enum_trailing_comma_typedef() { assert_eq!(run_c("typedef enum { A, B, } MyEnum; int main() { MyEnum e = B; printf(\"%d\", e); return 0; }"), vec!["1"]); }
#[test] fn enum_trailing_comma_in_struct() { assert_eq!(run_c("struct S { enum { A=10, } e; }; int main() { printf(\"%d\", A); return 0; }"), vec!["10"]); }
#[test] fn enum_trailing_comma_hex() { assert_eq!(run_c("enum E { A=0xA, B=0xB, }; int main() { printf(\"%d\", B); return 0; }"), vec!["11"]); }
#[test] fn enum_trailing_comma_negative() { assert_eq!(run_c("enum E { A=-1, B=-2, }; int main() { printf(\"%d\", B); return 0; }"), vec!["-2"]); }
#[test] fn enum_trailing_comma_multiple_commas_fails() { assert_eq!(run_c("/* enum E { A,, }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn enum_trailing_comma_no_elements_fails() { assert_eq!(run_c("/* enum E { , }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn enum_trailing_comma_maco_expansion() { assert_eq!(run_c("#define ELEMS A, B,\nenum E { ELEMS }; int main() { printf(\"%d\", B); return 0; }"), vec!["1"]); }
#[test] fn enum_trailing_comma_macro_end() { assert_eq!(run_c("#define COMMA ,\nenum E { A COMMA B COMMA }; int main() { printf(\"%d\", B); return 0; }"), vec!["1"]); }
#[test] fn enum_trailing_comma_with_sizeof() { assert_eq!(run_c("enum E { A, B, }; int main() { printf(\"%d\", (int)sizeof(enum E)); return 0; }"), vec!["4"]); }
#[test] fn enum_trailing_comma_in_func() { assert_eq!(run_c("int main() { enum { A, B, }; printf(\"%d\", B); return 0; }"), vec!["1"]); }
#[test] fn enum_trailing_comma_array_size() { assert_eq!(run_c("enum { SIZE=3, }; int main() { int arr[SIZE]; arr[2]=5; printf(\"%d\", arr[2]); return 0; }"), vec!["5"]); }
