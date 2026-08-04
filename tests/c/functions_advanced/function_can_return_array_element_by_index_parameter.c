// vybe-test: c/functions_advanced/function_can_return_array_element_by_index_parameter
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int get_at(int arr[], int index) { return arr[index]; }
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int arr[3] = {2, 4, 6};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", get_at(arr, 2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

