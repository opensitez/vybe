// vybe-test: c/c_vla_function_parameters/vla_parameter_typedef
// origin: languages/c/tests/c/test_c_vla_function_parameters.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 int n=5; typedef int VLA[n]; VLA a; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)(sizeof(a)/sizeof(int)));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

