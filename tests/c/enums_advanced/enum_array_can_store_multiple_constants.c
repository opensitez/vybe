// vybe-test: c/enums_advanced/enum_array_can_store_multiple_constants
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Digit { ZERO, ONE, TWO };
int main() {
const char *__w[] = {"0 1 2\n"};
int __n = 1, __i = 0;
enum Digit digits[3] = {ZERO, ONE, TWO}; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", digits[0], digits[1], digits[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

