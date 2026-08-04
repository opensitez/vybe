// vybe-test: c/lang_union_type_punning/union_postfix_increment_int_member
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; };
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
union U u; u.i = 1; int v = u.i++; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", v, u.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

