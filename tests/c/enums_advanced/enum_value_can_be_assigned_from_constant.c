// vybe-test: c/enums_advanced/enum_value_can_be_assigned_from_constant
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Mode { A = 4, B = 9 };
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
enum Mode mode = B; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", mode);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

