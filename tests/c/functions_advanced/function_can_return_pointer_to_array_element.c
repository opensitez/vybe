// vybe-test: c/functions_advanced/function_can_return_pointer_to_array_element
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int *second(int arr[]) { return &arr[1]; }
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
int values[3] = {8, 9, 10};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *second(values));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

