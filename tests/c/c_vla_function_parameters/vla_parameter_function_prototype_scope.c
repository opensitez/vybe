// vybe-test: c/c_vla_function_parameters/vla_parameter_function_prototype_scope
// origin: languages/c/tests/c/test_c_vla_function_parameters.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"2"};
static int __n = 1, __i = 0;
void f(int n, int a[n+1]); void f(int n, int a[n+1]) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { int arr[2] = {1, 2}; f(1, arr); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

