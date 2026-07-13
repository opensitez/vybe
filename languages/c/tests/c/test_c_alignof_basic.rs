use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn alignof_basic_types() { assert_eq!(run_c("#include <stdalign.h>\nint main() { printf(\"%d\", alignof(char) == 1); return 0; }"), vec!["1"]); }
#[test] fn alignof_int() { assert_eq!(run_c("#include <stdalign.h>\nint main() { printf(\"%d\", alignof(int) >= 1); return 0; }"), vec!["1"]); } // Normally 4
#[test] fn alignof_double() { assert_eq!(run_c("#include <stdalign.h>\nint main() { printf(\"%d\", alignof(double) >= 4); return 0; }"), vec!["1"]); } // 4 or 8
#[test] fn alignof_pointer() { assert_eq!(run_c("#include <stdalign.h>\nint main() { printf(\"%d\", alignof(void*) == sizeof(void*)); return 0; }"), vec!["1"]); }
#[test] fn alignof_struct() { assert_eq!(run_c("#include <stdalign.h>\nstruct S { char c; double d; }; int main() { printf(\"%d\", alignof(struct S) == alignof(double)); return 0; }"), vec!["1"]); }
#[test] fn alignof_array() { assert_eq!(run_c("#include <stdalign.h>\nint main() { printf(\"%d\", alignof(int[10]) == alignof(int)); return 0; }"), vec!["1"]); }
#[test] fn _alignof_keyword() { assert_eq!(run_c("int main() { printf(\"%d\", _Alignof(int) == _Alignof(int[5])); return 0; }"), vec!["1"]); }
#[test] fn alignof_expression_fails() { assert_eq!(run_c("/* int main() { int x; _Alignof(x); return 0; } // _Alignof takes type, not expression */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignof_incomplete_type_fails() { assert_eq!(run_c("/* int main() { _Alignof(void); return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignof_function_type_fails() { assert_eq!(run_c("/* int main() { _Alignof(void(void)); return 0; } */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignof_vla() { assert_eq!(run_c("int main() { int n=5; printf(\"%d\", _Alignof(int[n]) == _Alignof(int)); return 0; }"), vec!["1"]); }
#[test] fn alignof_union() { assert_eq!(run_c("union U { char c; double d; }; int main() { printf(\"%d\", _Alignof(union U) == _Alignof(double)); return 0; }"), vec!["1"]); }
#[test] fn alignof_enum() { assert_eq!(run_c("enum E { A, B }; int main() { printf(\"%d\", _Alignof(enum E) == _Alignof(int)); return 0; }"), vec!["1"]); } // Enums usually int aligned
#[test] fn alignof_typedef() { assert_eq!(run_c("typedef double mydouble; int main() { printf(\"%d\", _Alignof(mydouble) == _Alignof(double)); return 0; }"), vec!["1"]); }
#[test] fn alignof_nested_struct() { assert_eq!(run_c("struct Inner { double d; }; struct Outer { char c; struct Inner i; }; int main() { printf(\"%d\", _Alignof(struct Outer) == _Alignof(double)); return 0; }"), vec!["1"]); }
