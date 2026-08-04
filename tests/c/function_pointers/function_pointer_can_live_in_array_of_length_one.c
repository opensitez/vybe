// vybe-test: c/function_pointers/function_pointer_can_live_in_array_of_length_one
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int negate(int x) { return -x; }
int main() {
const char *__w[] = {"-5\n"};
int __n = 1, __i = 0;
int (*ops[1])(int) = {negate};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ops[0](5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

