// vybe-test: c/c_pointer_void_cast_arithmetic/void_pointer_cast_to_char
// origin: languages/c/tests/c/test_c_pointer_void_cast_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int x = 0x12345678; void *p = &x; char *cp = p; /* accesses byte, endian dependent, just test compile */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

