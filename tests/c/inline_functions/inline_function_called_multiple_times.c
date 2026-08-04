// vybe-test: c/inline_functions/inline_function_called_multiple_times
// origin: languages/c/tests/c/test_inline_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static inline int add_one(int x) { return x + 1; }
int main() {
const char *__w[] = {"2 11 101\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", add_one(1), add_one(10), add_one(100));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

