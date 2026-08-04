// vybe-test: c/c_pointer_to_function_arrays/pointer_func_array_typedef
// origin: languages/c/tests/c/test_c_pointer_to_function_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int (*F)(void); int f(){return 3;} int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 F arr[1] = {f}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", arr[0]());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

