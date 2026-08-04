// vybe-test: c/lang_union_type_punning/union_struct_anonymous_in_union_outer
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { struct { int x; int y; } s; int i; };
int main() {
const char *__w[] = {"6 7\n"};
int __n = 1, __i = 0;
union U u; u.s.x = 6; u.s.y = 7; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", u.s.x, u.s.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

