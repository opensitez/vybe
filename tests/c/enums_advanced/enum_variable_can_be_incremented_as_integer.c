// vybe-test: c/enums_advanced/enum_variable_can_be_incremented_as_integer
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Count { ZERO, ONE, TWO };
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
enum Count count = ZERO; count = count + 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

