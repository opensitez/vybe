// vybe-test: c/nested_functions/function_taking_array_and_size
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int array_sum(int *arr, int n) { int s = 0; for (int i = 0; i < n; i++) s += arr[i]; return s; }
int main() {
const char *__w[] = {"15\n"};
int __n = 1, __i = 0;
int data[] = {1,2,3,4,5};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", array_sum(data, 5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

