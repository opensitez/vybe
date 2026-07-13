use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn type_punning_basic() { assert_eq!(run_c("union U { int i; float f; }; int main() { union U u; u.f = 3.14f; /* C99 standard permits reading inactive union member */ int val = u.i; printf(\"%d\", val != 0); return 0; }"), vec!["1"]); }
#[test] fn type_punning_char_array() { assert_eq!(run_c("union U { int i; char c[4]; }; int main() { union U u; u.i = 0x41424344; printf(\"%d\", u.c[0] != 0); return 0; }"), vec!["1"]); }
#[test] fn type_punning_structs() { assert_eq!(run_c("struct A { int type; int val; }; struct B { int type; float f; }; union U { struct A a; struct B b; }; int main() { union U u; u.a.type = 1; u.a.val = 5; printf(\"%d\", u.b.type); return 0; }"), vec!["1"]); } // Valid in C99 if initial sequence matches
#[test] fn type_punning_pointer_fails() { assert_eq!(run_c("int main() { float f = 3.14f; /* int *p = (int*)&f; // Strict aliasing violation */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn type_punning_memcpy() { assert_eq!(run_c("#include <string.h>\nint main() { float f = 3.14f; int i; memcpy(&i, &f, sizeof(float)); printf(\"%d\", i != 0); return 0; }"), vec!["1"]); } // The safe way to pun
#[test] fn type_punning_union_pointers() { assert_eq!(run_c("union U { int *p_i; char *p_c; }; int main() { int val = 65; union U u; u.p_i = &val; printf(\"%c\", *u.p_c); return 0; }"), vec!["A"]); } // Typically works on LE
#[test] fn type_punning_bitfields() { assert_eq!(run_c("union U { int i; struct { unsigned int a:4; unsigned int b:4; } b; }; int main() { union U u; u.i = 0xFF; printf(\"%d\", u.b.a); return 0; }"), vec!["15"]); }
#[test] fn type_punning_array_to_struct() { assert_eq!(run_c("union U { int arr[2]; struct { int x; int y; } s; }; int main() { union U u; u.arr[0] = 1; u.arr[1] = 2; printf(\"%d\", u.s.y); return 0; }"), vec!["2"]); }
#[test] fn type_punning_struct_to_array() { assert_eq!(run_c("union U { struct { int x; int y; } s; int arr[2]; }; int main() { union U u; u.s.x = 10; u.s.y = 20; printf(\"%d\", u.arr[1]); return 0; }"), vec!["20"]); }
#[test] fn type_punning_union_of_unions() { assert_eq!(run_c("union A { int x; }; union B { float y; }; union U { union A a; union B b; }; int main() { union U u; u.a.x = 100; printf(\"%d\", u.b.y != 0.0); return 0; }"), vec!["1"]); }
#[test] fn type_punning_char_ptr_cast() { assert_eq!(run_c("int main() { int x = 0x12345678; char *p = (char*)&x; printf(\"%d\", *p != 0); return 0; }"), vec!["1"]); } // char* is allowed to alias any type
#[test] fn type_punning_void_ptr() { assert_eq!(run_c("int main() { int x = 42; void *p = &x; float *fp = (float*)p; /* Valid to cast, but deref is UB. We just test compilation */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn type_punning_union_initialization() { assert_eq!(run_c("union U { int i; char c; }; int main() { union U u = {65}; printf(\"%c\", u.c); return 0; }"), vec!["A"]); }
#[test] fn type_punning_union_assignment() { assert_eq!(run_c("union U { int i; float f; }; int main() { union U u1, u2; u1.f = 2.5f; u2 = u1; printf(\"%d\", u2.f > 2.0); return 0; }"), vec!["1"]); }
#[test] fn type_punning_struct_padding() { assert_eq!(run_c("union U { struct { char c; int i; } s; char arr[8]; }; int main() { union U u; u.s.c = 'A'; u.s.i = 5; /* padding bytes are indeterminate */ printf(\"ok\"); return 0; }"), vec!["ok"]); }
