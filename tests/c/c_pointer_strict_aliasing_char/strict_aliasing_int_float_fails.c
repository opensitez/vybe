// vybe-test: c/c_pointer_strict_aliasing_char/strict_aliasing_int_float_fails
// origin: languages/c/tests/c/test_c_pointer_strict_aliasing_char.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* int main() { float f = 1.0f; int *p = (int*)&f; *p = 1; return 0; } // UB */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

