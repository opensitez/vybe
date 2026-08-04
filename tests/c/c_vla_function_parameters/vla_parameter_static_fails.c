// vybe-test: c/c_vla_function_parameters/vla_parameter_static_fails
// origin: languages/c/tests/c/test_c_vla_function_parameters.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* void f(int n, int arr[static n]) {} // static requires at least n elements */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

