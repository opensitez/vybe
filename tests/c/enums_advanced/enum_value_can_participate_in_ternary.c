// vybe-test: c/enums_advanced/enum_value_can_participate_in_ternary
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Light { RED, GREEN };
int main() {
const char *__w[] = {"go\n"};
int __n = 1, __i = 0;
enum Light light = GREEN; { char __t[512]; snprintf(__t, sizeof(__t), "%s\n", light == GREEN ? "go" : "stop");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

