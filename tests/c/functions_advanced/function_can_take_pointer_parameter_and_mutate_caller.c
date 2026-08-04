// vybe-test: c/functions_advanced/function_can_take_pointer_parameter_and_mutate_caller
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void set_to_ten(int *p) { *p = 10; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int value = 3;
set_to_ten(&value);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", value);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

