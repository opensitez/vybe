// vybe-test: c/function_pointers/function_pointer_parameter_can_use_local_choice
// origin: languages/c/tests/c/test_function_pointers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int add_one(int x) { return x + 1; }
int double_it(int x) { return x * 2; }
int call_selected(int which, int value) { int (*fp)(int) = which ? double_it : add_one; return fp(value); }
int main() {
const char *__w[] = {"4 6\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", call_selected(0, 3), call_selected(1, 3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

