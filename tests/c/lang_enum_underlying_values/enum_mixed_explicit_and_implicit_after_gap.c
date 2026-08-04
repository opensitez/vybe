// vybe-test: c/lang_enum_underlying_values/enum_mixed_explicit_and_implicit_after_gap
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { A = 1, B = 10, C, D = 20, E };
int main() {
const char *__w[] = {"1 10 11 20 21\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d %d\n", A, B, C, D, E);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

