// vybe-test: c/c_comma_operator_function_args/comma_operator_in_array_init
// origin: languages/c/tests/c/test_c_comma_operator_function_args.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"24"};
int __n = 1, __i = 0;
 int arr[] = { (1, 2), (3, 4) }; { char __t[512]; snprintf(__t, sizeof(__t), "%d%d", arr[0], arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

