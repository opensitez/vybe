// vybe-test: c/printf_length_modifiers/printf_lf_long_double_literal
// origin: languages/c/tests/c/test_printf_length_modifiers.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stddef.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"1.250000\n"};
int __n = 1, __i = 0;
long double ld=1.25L; { char __t[512]; snprintf(__t, sizeof(__t), "%Lf\n", ld);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

