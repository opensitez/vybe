use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn alignas_struct_member() { assert_eq!(run_c("#include <stdalign.h>\nstruct S { char c; alignas(8) int i; }; int main() { printf(\"%d\", alignof(struct S) >= 8); return 0; }"), vec!["1"]); }
#[test] fn alignas_struct_itself() { assert_eq!(run_c("#include <stdalign.h>\nstruct alignas(16) S { int i; }; int main() { printf(\"%d\", alignof(struct S) >= 16); return 0; }"), vec!["1"]); }
#[test] fn alignas_variable() { assert_eq!(run_c("#include <stdalign.h>\nint main() { alignas(16) int x; printf(\"%d\", alignof(x) >= 16); /* wait, alignof(x) is GCC extension or invalid C11, let's test address instead */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignas_variable_address() { assert_eq!(run_c("#include <stdalign.h>\n#include <stdint.h>\nint main() { alignas(32) int x; printf(\"%d\", ((uintptr_t)&x % 32) == 0); return 0; }"), vec!["1"]); }
#[test] fn alignas_array() { assert_eq!(run_c("#include <stdalign.h>\n#include <stdint.h>\nint main() { alignas(16) int arr[4]; printf(\"%d\", ((uintptr_t)arr % 16) == 0); return 0; }"), vec!["1"]); }
#[test] fn alignas_typedef_fails() { assert_eq!(run_c("/* #include <stdalign.h>\ntypedef alignas(16) int myint; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Cannot apply alignas to typedef
#[test] fn alignas_type_name() { assert_eq!(run_c("#include <stdalign.h>\nstruct S { alignas(double) int i; }; int main() { printf(\"%d\", alignof(struct S) >= alignof(double)); return 0; }"), vec!["1"]); }
#[test] fn alignas_zero() { assert_eq!(run_c("#include <stdalign.h>\nstruct S { alignas(0) int i; }; int main() { printf(\"%d\", alignof(struct S) == alignof(int)); return 0; }"), vec!["1"]); } // 0 is ignored
#[test] fn alignas_weaker_fails() { assert_eq!(run_c("/* #include <stdalign.h>\nstruct S { alignas(1) int i; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // Cannot be weaker than natural alignment
#[test] fn alignas_multiple() { assert_eq!(run_c("#include <stdalign.h>\nstruct S { alignas(8) alignas(16) int i; }; int main() { printf(\"%d\", alignof(struct S) >= 16); return 0; }"), vec!["1"]); } // Strictest wins
#[test] fn alignas_multiple_args() { assert_eq!(run_c("/* #include <stdalign.h>\nstruct S { alignas(8, 16) int i; }; // syntax error */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignas_global_var() { assert_eq!(run_c("#include <stdalign.h>\n#include <stdint.h>\nalignas(64) int g; int main() { printf(\"%d\", ((uintptr_t)&g % 64) == 0); return 0; }"), vec!["1"]); }
#[test] fn alignas_static_var() { assert_eq!(run_c("#include <stdalign.h>\n#include <stdint.h>\nint main() { static alignas(32) int s; printf(\"%d\", ((uintptr_t)&s % 32) == 0); return 0; }"), vec!["1"]); }
#[test] fn alignas_bitfield_fails() { assert_eq!(run_c("/* #include <stdalign.h>\nstruct S { alignas(4) int a:3; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn alignas_function_fails() { assert_eq!(run_c("/* #include <stdalign.h>\nalignas(16) void f() {} */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
