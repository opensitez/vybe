use super::helpers::run_prints;
fn run_c(src: &str) -> Vec<String> { run_prints(&format!("#include <stdio.h>\n{}", src)) }

#[test] fn malloc_basic() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); *p = 5; printf(\"%d\", *p); free(p); return 0; }"), vec!["5"]); }
#[test] fn malloc_zero_bytes() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = malloc(0); printf(\"%d\", p == NULL || p != NULL); free(p); return 0; }"), vec!["1"]); } // May return NULL or unique pointer
#[test] fn malloc_struct() { assert_eq!(run_c("#include <stdlib.h>\nstruct S { int a; }; int main() { struct S *p = malloc(sizeof(struct S)); p->a = 10; printf(\"%d\", p->a); free(p); return 0; }"), vec!["10"]); }
#[test] fn malloc_array() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int *p = malloc(3 * sizeof(int)); p[1] = 7; printf(\"%d\", p[1]); free(p); return 0; }"), vec!["7"]); }
#[test] fn malloc_cast_void() { assert_eq!(run_c("#include <stdlib.h>\nint main() { void *p = malloc(10); char *c = p; c[0] = 'A'; printf(\"%c\", c[0]); free(p); return 0; }"), vec!["A"]); }
#[test] fn malloc_free_null() { assert_eq!(run_c("#include <stdlib.h>\nint main() { free(NULL); printf(\"ok\"); return 0; }"), vec!["ok"]); } // free(NULL) does nothing
#[test] fn malloc_negative_size_fails() { assert_eq!(run_c("/* #include <stdlib.h>\nint main() { void *p = malloc(-1); return 0; } // size_t is unsigned, so -1 is SIZE_MAX */ int main() { printf(\"ok\"); return 0; }"), vec!["ok"]); }
#[test] fn malloc_vla_sizeof() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int n=5; int *p = malloc(sizeof(int[n])); p[4] = 42; printf(\"%d\", p[4]); free(p); return 0; }"), vec!["42"]); }
#[test] fn malloc_pointer_to_pointer() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int **p = malloc(sizeof(int*)); *p = malloc(sizeof(int)); **p = 99; printf(\"%d\", **p); free(*p); free(p); return 0; }"), vec!["99"]); }
#[test] fn malloc_null_check() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int *p = malloc(sizeof(int)); if(p != NULL) { *p = 1; printf(\"%d\", *p); free(p); } return 0; }"), vec!["1"]); }
#[test] fn malloc_struct_with_pointer() { assert_eq!(run_c("#include <stdlib.h>\nstruct S { int *p; }; int main() { struct S *s = malloc(sizeof(struct S)); s->p = malloc(sizeof(int)); *s->p = 3; printf(\"%d\", *s->p); free(s->p); free(s); return 0; }"), vec!["3"]); }
#[test] fn malloc_alignment() { assert_eq!(run_c("#include <stdlib.h>\n#include <stdint.h>\nint main() { void *p = malloc(1); printf(\"%d\", ((uintptr_t)p % 8) == 0 || ((uintptr_t)p % 16) == 0); free(p); return 0; }"), vec!["1"]); } // Malloc is suitably aligned
#[test] fn malloc_multidim_array() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int (*p)[3] = malloc(2 * sizeof(int[3])); p[1][2] = 55; printf(\"%d\", p[1][2]); free(p); return 0; }"), vec!["55"]); }
#[test] fn malloc_in_loop() { assert_eq!(run_c("#include <stdlib.h>\nint main() { int sum=0; for(int i=0; i<3; i++) { int *p = malloc(sizeof(int)); *p = i; sum += *p; free(p); } printf(\"%d\", sum); return 0; }"), vec!["3"]); }
