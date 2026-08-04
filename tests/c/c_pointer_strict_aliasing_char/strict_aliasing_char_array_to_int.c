// vybe-test: c/c_pointer_strict_aliasing_char/strict_aliasing_char_array_to_int
// origin: languages/c/tests/c/test_c_pointer_strict_aliasing_char.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 char arr[sizeof(int)] = {0}; int *p = (int*)arr; /* UB if unaligned, or strict aliasing violation if arr is declared char. But often done. */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

