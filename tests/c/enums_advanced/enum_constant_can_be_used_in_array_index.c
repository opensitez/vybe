// vybe-test: c/enums_advanced/enum_constant_can_be_used_in_array_index
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Slot { FIRST, SECOND, THIRD };
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;
int values[3] = {10, 20, 30}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", values[SECOND]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

