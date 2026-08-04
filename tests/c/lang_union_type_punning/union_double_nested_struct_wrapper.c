// vybe-test: c/lang_union_type_punning/union_double_nested_struct_wrapper
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; }; struct Mid { union U u; }; struct Top { struct Mid m; };
int main() {
const char *__w[] = {"81\n"};
int __n = 1, __i = 0;
struct Top t; t.m.u.i = 81; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", t.m.u.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

