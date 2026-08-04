// vybe-test: c/enums_advanced/enum_typedef_and_variable_can_share_namespace_rules
// origin: languages/c/tests/c/test_enums_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef enum { APPLE, PEAR } Fruit;
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
Fruit fruit = PEAR; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", fruit);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

