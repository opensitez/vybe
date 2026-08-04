// vybe-test: c/lang_enum_underlying_values/enum_mixed_add_int_and_enum
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum B { BASE = 5 };
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
int x = BASE + 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

