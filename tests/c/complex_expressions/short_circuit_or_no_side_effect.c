// vybe-test: c/complex_expressions/short_circuit_or_no_side_effect
// origin: languages/c/tests/c/test_complex_expressions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"5 0\n"};
int __n = 1, __i = 0;
int x = 0;
int y = 1;
(y = 0) || (x = 5);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", x, y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

