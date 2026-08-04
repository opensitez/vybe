// vybe-test: c/ternary_comma/ternary_and_comma_can_combine_in_parenthesized_expression
// origin: languages/c/tests/c/test_ternary_comma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int x = 1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (x = 0, x ? 1 : 2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

