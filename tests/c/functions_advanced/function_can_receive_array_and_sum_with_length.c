// vybe-test: c/functions_advanced/function_can_receive_array_and_sum_with_length
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int sum(int arr[], int len) { int total = 0; for (int i = 0; i < len; i++) total += arr[i]; return total; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int data[4] = {1, 2, 3, 4};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum(data, 4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

