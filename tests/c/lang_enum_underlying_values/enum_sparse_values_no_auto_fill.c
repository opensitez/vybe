// vybe-test: c/lang_enum_underlying_values/enum_sparse_values_no_auto_fill
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum E { X = 10, Y = 20, Z = 30 };
int main() {
const char *__w[] = {"10 20 30\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", X, Y, Z);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

