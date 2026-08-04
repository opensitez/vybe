// vybe-test: c/c_pointer_arithmetic_past_end/pointer_past_end_struct_array
// origin: languages/c/tests/c/test_c_pointer_arithmetic_past_end.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int x; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct S arr[2]; struct S *p = arr + 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p > arr);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

