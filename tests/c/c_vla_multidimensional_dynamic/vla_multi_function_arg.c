// vybe-test: c/c_vla_multidimensional_dynamic/vla_multi_function_arg
// origin: languages/c/tests/c/test_c_vla_multidimensional_dynamic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"1"};
static int __n = 1, __i = 0;
void f(int r, int c, int arr[r][c]) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)sizeof(*arr) == c * sizeof(int));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { int r=2, c=3; int arr[r][c]; f(r, c, arr); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

