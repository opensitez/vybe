// vybe-test: c/c_pointer_strict_aliasing_char/strict_aliasing_compatible_types
// origin: languages/c/tests/c/test_c_pointer_strict_aliasing_char.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 signed int x = 5; unsigned int *p = (unsigned int*)&x; { char __t[512]; snprintf(__t, sizeof(__t), "%u", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

