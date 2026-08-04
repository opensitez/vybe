// vybe-test: c/sizeof/sizeof_function_parameter_array_is_pointer_size
// origin: languages/c/tests/c/test_sizeof.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int size_of_param(int arr[]) { return (int)sizeof(arr); }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
int arr[3] = {0}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", size_of_param(arr));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

