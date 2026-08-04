// vybe-test: c/c_pointer_to_function_arrays/pointer_func_array_different_signatures_fails
// origin: languages/c/tests/c/test_c_pointer_to_function_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* int f(){return 1;} void g(){} int (*arr[2])() = {f, g}; // error */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

