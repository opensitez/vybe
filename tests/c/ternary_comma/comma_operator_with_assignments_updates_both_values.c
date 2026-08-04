// vybe-test: c/ternary_comma/comma_operator_with_assignments_updates_both_values
// origin: languages/c/tests/c/test_ternary_comma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"2 3\n"};
int __n = 1, __i = 0;
int a = 0; int b = 0; (a = 2, b = 3); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", a, b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

