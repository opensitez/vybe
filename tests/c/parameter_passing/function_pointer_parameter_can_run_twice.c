// vybe-test: c/parameter_passing/function_pointer_parameter_can_run_twice
// origin: languages/c/tests/c/test_parameter_passing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int apply_twice(int (*fn)(int), int value) { return fn(fn(value)); } int inc(int x) { return x + 1; }
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", apply_twice(inc, 3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

