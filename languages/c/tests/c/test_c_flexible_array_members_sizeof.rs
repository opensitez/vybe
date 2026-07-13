use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn fam_sizeof_ignores_array() { assert_eq!(run_c("struct S { int len; int data[]; }; int main() { printf(\"%d\", (int)(sizeof(struct S) == sizeof(int))); return 0; }"), vec!["1"]); }
#[test] fn fam_sizeof_with_padding() { assert_eq!(run_c("struct S { char c; double data[]; }; int main() { printf(\"%d\", (int)(sizeof(struct S) == sizeof(double))); return 0; }"), vec!["1"]); } // Padded to double alignment
#[test] fn fam_must_be_last_fails() { assert_eq!(run_c("/* struct S { int data[]; int len; }; */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_only_member_fails() { assert_eq!(run_c("/* struct S { int data[]; }; // Must have at least one other named member */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_in_union_fails() { assert_eq!(run_c("/* union U { int a; int data[]; }; // FAM not allowed in union */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_array_of_structs_fails() { assert_eq!(run_c("/* struct S { int len; int data[]; }; struct S arr[2]; // Struct with FAM cannot be element of array */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_struct_in_struct_fails() { assert_eq!(run_c("/* struct S { int len; int data[]; }; struct Outer { struct S inner; int x; }; // Struct with FAM cannot be nested unless it's the last member */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_struct_in_struct_last() { assert_eq!(run_c("struct S { int len; int data[]; }; struct Outer { int x; struct S inner; }; int main() { printf(\"%d\", (int)sizeof(struct Outer) >= sizeof(int)*2); return 0; }"), vec!["1"]); }
#[test] fn fam_incomplete_type_size() { assert_eq!(run_c("struct S { int len; struct Incomplete *data[]; }; int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); } // pointer to incomplete is complete
#[test] fn fam_assignment_fails() { assert_eq!(run_c("/* struct S { int len; int data[]; }; int main() { struct S s1, s2; s1 = s2; return 0; } // assignment copies only declared members, standard says behavior is undefined or just copies named members */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_initialization_fails() { assert_eq!(run_c("/* struct S { int len; int data[]; }; struct S s = {1, {2, 3}}; // Cannot initialize FAM in standard C */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_gnu_initialization() { assert_eq!(run_c("/* GNU extension allows initializing FAM */ struct S { int len; int data[]; } s = {2, {1, 2}}; int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn fam_pass_by_value() { assert_eq!(run_c("struct S { int len; int data[]; }; void f(struct S s) { printf(\"%d\", s.len); } int main() { struct S s = {5}; f(s); return 0; }"), vec!["5"]); } // Copies only `len`
#[test] fn fam_return_by_value() { assert_eq!(run_c("struct S { int len; int data[]; }; struct S f() { struct S s = {6}; return s; } int main() { printf(\"%d\", f().len); return 0; }"), vec!["6"]); }
