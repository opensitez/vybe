// vybe-test: c/c_vla_multidimensional_dynamic/vla_multi_typedef
// origin: languages/c/tests/c/test_c_vla_multidimensional_dynamic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"7"};
int __n = 1, __i = 0;
 int c=3; typedef int Row[c]; Row arr[2]; arr[1][2] = 7; { char __t[512]; snprintf(__t, sizeof(__t), "%d", arr[1][2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

