// vybe-test: c/lang_enum_underlying_values/enum_compare_after_cast_from_int
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { A = 0, B = 1 };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
enum E e = (enum E)1; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", e == B);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

