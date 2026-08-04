// vybe-test: c/lang_enum_underlying_values/enum_ternary_with_enum_condition
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum B { NO, YES };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
enum B b = YES; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", b ? 1 : 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

