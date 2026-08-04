// vybe-test: c/enums_advanced/enum_expression_can_feed_array_size_like_constant_usage
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum Size { LEN = 3 };
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
int values[LEN] = {1, 2, 3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", values[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

