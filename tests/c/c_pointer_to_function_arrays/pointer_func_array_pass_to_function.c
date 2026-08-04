// vybe-test: c/c_pointer_to_function_arrays/pointer_func_array_pass_to_function
// origin: languages/c/tests/c/test_c_pointer_to_function_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"5"};
static int __n = 1, __i = 0;
int f(){return 5;} void run(int (*a[])()) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", a[0]());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { int (*arr[1])() = {f}; run(arr); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

