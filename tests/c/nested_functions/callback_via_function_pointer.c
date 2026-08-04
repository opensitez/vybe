// vybe-test: c/nested_functions/callback_via_function_pointer
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int apply(int x, int (*fn)(int)) { return fn(x); }
int triple(int x) { return x * 3; }
int main() {
const char *__w[] = {"15\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", apply(5, triple));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

