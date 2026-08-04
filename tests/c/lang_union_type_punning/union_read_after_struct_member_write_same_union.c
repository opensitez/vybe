// vybe-test: c/lang_union_type_punning/union_read_after_struct_member_write_same_union
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; short s; }; struct S { union U u; };
int main() {
const char *__w[] = {"77\n"};
int __n = 1, __i = 0;
struct S s; s.u.i = 77; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s.u.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

