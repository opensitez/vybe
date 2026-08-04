// vybe-test: c/c_pointer_to_array_multidim/pointer_multidim_return_from_function
// origin: languages/c/tests/c/test_c_pointer_to_array_multidim.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[2][2] = {{1,2},{3,4}}; int (*f())[2] { return arr; } int main() {const char *__w[] = {"4"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", f()[1][1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

