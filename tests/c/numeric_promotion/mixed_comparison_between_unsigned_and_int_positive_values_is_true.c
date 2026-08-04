// vybe-test: c/numeric_promotion/mixed_comparison_between_unsigned_and_int_positive_values_is_true
// origin: languages/c/tests/c/test_numeric_promotion.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
unsigned int u = 5; int i = 5; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", u == i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

