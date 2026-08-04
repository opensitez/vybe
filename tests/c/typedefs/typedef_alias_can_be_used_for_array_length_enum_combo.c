// vybe-test: c/typedefs/typedef_alias_can_be_used_for_array_length_enum_combo
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Number; enum { LEN = 3 };
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
Number values[LEN] = {2, 4, 6}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", values[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

