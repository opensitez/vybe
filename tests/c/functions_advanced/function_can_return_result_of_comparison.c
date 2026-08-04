// vybe-test: c/functions_advanced/function_can_return_result_of_comparison
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int is_even(int x) { return x % 2 == 0; }
int main() {
const char *__w[] = {"1 0\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", is_even(4), is_even(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

