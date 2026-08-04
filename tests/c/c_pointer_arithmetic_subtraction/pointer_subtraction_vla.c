// vybe-test: c/c_pointer_arithmetic_subtraction/pointer_subtraction_vla
// origin: languages/c/tests/c/test_c_pointer_arithmetic_subtraction.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 int n=5; int arr[n]; int *p1 = &arr[1]; int *p2 = &arr[n-1]; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)(p2 - p1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

