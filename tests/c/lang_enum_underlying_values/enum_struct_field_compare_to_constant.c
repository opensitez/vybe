// vybe-test: c/lang_enum_underlying_values/enum_struct_field_compare_to_constant
// origin: languages/c/tests/c/test_lang_enum_underlying_values.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
enum L { LOW, MID, HIGH }; struct S { enum L level; };
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
struct S s = {HIGH}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", s.level == HIGH);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

