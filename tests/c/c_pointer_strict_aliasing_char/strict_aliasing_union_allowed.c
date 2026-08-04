// vybe-test: c/c_pointer_strict_aliasing_char/strict_aliasing_union_allowed
// origin: languages/c/tests/c/test_c_pointer_strict_aliasing_char.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union U { int i; float f; }; int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 union U u; u.f = 1.0f; int val = u.i; { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

