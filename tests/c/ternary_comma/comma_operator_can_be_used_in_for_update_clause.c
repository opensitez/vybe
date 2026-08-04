// vybe-test: c/ternary_comma/comma_operator_can_be_used_in_for_update_clause
// origin: languages/c/tests/c/test_ternary_comma.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"02\n", "11\n"};
int __n = 2, __i = 0;
for (int i = 0, j = 2; i < 2; i++, j--) { char __t[512]; snprintf(__t, sizeof(__t), "%d%d\n", i, j);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

