// vybe-test: c/c_sizeof_vla/sizeof_vla_2d_partial
// origin: languages/c/tests/c/test_c_sizeof_vla.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 int n = 2, m = 3; int arr[n][m]; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)(sizeof(arr[0]) / sizeof(int)));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

