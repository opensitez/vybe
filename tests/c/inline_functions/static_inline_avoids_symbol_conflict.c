// vybe-test: c/inline_functions/static_inline_avoids_symbol_conflict
// origin: languages/c/tests/c/test_inline_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static inline double halve(double x) { return x / 2.0; }
int main() {
const char *__w[] = {"3.5\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", halve(7.0));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

