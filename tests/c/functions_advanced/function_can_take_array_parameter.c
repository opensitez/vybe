// vybe-test: c/functions_advanced/function_can_take_array_parameter
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int first(int arr[]) { return arr[0]; }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
int values[3] = {8, 9, 10};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", first(values));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

