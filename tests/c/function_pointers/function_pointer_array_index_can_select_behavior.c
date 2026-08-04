// vybe-test: c/function_pointers/function_pointer_array_index_can_select_behavior
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int add_one(int x) { return x + 1; }
int double_it(int x) { return x * 2; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int (*ops[2])(int) = {add_one, double_it}; int index = 1;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", ops[index](5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

