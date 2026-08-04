// vybe-test: c/functions_advanced/function_can_mutate_array_element_through_parameter
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void set_second(int arr[]) { arr[1] = 99; }
int main() {
const char *__w[] = {"99\n"};
int __n = 1, __i = 0;
int arr[3] = {1, 2, 3};
set_second(arr);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

