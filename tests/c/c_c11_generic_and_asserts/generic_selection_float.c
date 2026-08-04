// vybe-test: c/c_c11_generic_and_asserts/generic_selection_float
// origin: languages/c/tests/c/test_c_c11_generic_and_asserts.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 int type_id = _Generic(3.14f, int: 1, float: 2, default: 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", type_id);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

