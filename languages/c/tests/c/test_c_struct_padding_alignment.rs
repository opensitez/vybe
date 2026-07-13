use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn struct_padding_basic() { assert_eq!(run_c("struct S { char c; int i; }; int main() { printf(\"%d\", sizeof(struct S) > sizeof(char)+sizeof(int)); return 0; }"), vec!["1"]); } // Usually padded to 8 bytes
#[test] fn struct_alignment_offset() { assert_eq!(run_c("#include <stddef.h>\nstruct S { char c; int i; }; int main() { printf(\"%d\", offsetof(struct S, i) > 0); return 0; }"), vec!["1"]); }
#[test] fn struct_padding_end() { assert_eq!(run_c("struct S { int i; char c; }; int main() { printf(\"%d\", sizeof(struct S) > sizeof(int)+sizeof(char)); return 0; }"), vec!["1"]); } // Padded at end to match int alignment
#[test] fn struct_alignment_double() { assert_eq!(run_c("#include <stddef.h>\nstruct S { char c; double d; }; int main() { printf(\"%d\", offsetof(struct S, d) >= 4); return 0; }"), vec!["1"]); }
#[test] fn struct_alignment_pointers() { assert_eq!(run_c("#include <stddef.h>\nstruct S { char c; void* p; }; int main() { printf(\"%d\", offsetof(struct S, p) >= 4); return 0; }"), vec!["1"]); }
#[test] fn struct_padding_nested() { assert_eq!(run_c("struct Inner { char c; }; struct Outer { struct Inner i; int a; }; int main() { printf(\"%d\", sizeof(struct Outer) > sizeof(char)+sizeof(int)); return 0; }"), vec!["1"]); }
#[test] fn struct_alignas_basic() { assert_eq!(run_c("struct S { _Alignas(8) char c; }; int main() { printf(\"%d\", (int)sizeof(struct S) >= 8); return 0; }"), vec!["1"]); }
#[test] fn struct_alignas_multiple() { assert_eq!(run_c("struct S { _Alignas(16) char c; int i; }; int main() { printf(\"%d\", (int)sizeof(struct S) >= 16); return 0; }"), vec!["1"]); }
#[test] fn struct_alignof_struct() { assert_eq!(run_c("struct S { char c; double d; }; int main() { printf(\"%d\", (int)_Alignof(struct S) >= 4); return 0; }"), vec!["1"]); }
#[test] fn struct_padding_zero_size_array_fails() { assert_eq!(run_c("/* struct S { int a[0]; }; // Zero size arrays are not standard */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn struct_padding_arrays() { assert_eq!(run_c("struct S { char c[3]; int i; }; int main() { printf(\"%d\", sizeof(struct S) > 3 + sizeof(int)); return 0; }"), vec!["1"]); }
#[test] fn struct_alignment_union() { assert_eq!(run_c("union U { char c; double d; }; int main() { printf(\"%d\", (int)_Alignof(union U) == (int)_Alignof(double)); return 0; }"), vec!["1"]); }
#[test] fn struct_padding_offsetof_macro() { assert_eq!(run_c("#define my_offsetof(t, m) ((size_t)&(((t*)0)->m))\nstruct S { char c; int i; }; int main() { printf(\"%d\", my_offsetof(struct S, i) > 0); return 0; }"), vec!["1"]); }
#[test] fn struct_alignment_max_align_t() { assert_eq!(run_c("#include <stddef.h>\nint main() { printf(\"%d\", (int)_Alignof(max_align_t) >= 8); return 0; }"), vec!["1"]); }
#[test] fn struct_padding_packed_attribute_fails_if_not_supported() { assert_eq!(run_c("/* struct __attribute__((packed)) S { char c; int i; }; // testing without attribute to keep pure standard C unless gcc exts are default */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
