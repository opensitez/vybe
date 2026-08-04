// vybe-test: c/c_c11_generic_and_asserts/generic_selection_array_decay
// origin: languages/c/tests/c/test_c_c11_generic_and_asserts.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int arr[5]; int type_id = _Generic(arr, int*: 1, int[5]: 2, default: 0); /* Arrays decay to pointers in _Generic */ { char __t[512]; snprintf(__t, sizeof(__t), "%d", type_id);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

