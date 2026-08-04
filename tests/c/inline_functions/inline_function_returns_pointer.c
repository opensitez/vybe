// vybe-test: c/inline_functions/inline_function_returns_pointer
// origin: languages/c/tests/c/test_inline_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static inline const char* greet(int n) { return n > 0 ? "positive" : "nonpositive"; }
int main() {
const char *__w[] = {"positive nonpositive\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %s\n", greet(5), greet(-3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

