// vybe-test: c/preprocessor/macro_can_expand_to_cast_expression
// origin: languages/c/tests/c/test_preprocessor.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define HALF(x) ((double)(x) / 2.0)
int main() {
const char *__w[] = {"3.5\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%.1f\n", HALF(7));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

