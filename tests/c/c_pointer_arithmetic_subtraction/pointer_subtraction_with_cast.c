// vybe-test: c/c_pointer_arithmetic_subtraction/pointer_subtraction_with_cast
// origin: languages/c/tests/c/test_c_pointer_arithmetic_subtraction.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int arr[2]; char *p1 = (char*)&arr[0]; char *p2 = (char*)&arr[1]; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)(p2 - p1) == sizeof(int));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

