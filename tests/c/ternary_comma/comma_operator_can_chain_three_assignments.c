// vybe-test: c/ternary_comma/comma_operator_can_chain_three_assignments
// origin: languages/c/tests/c/test_ternary_comma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int a = 0; int b = 0; int c = 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (a = 1, b = 2, c = 3, a + b + c));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

