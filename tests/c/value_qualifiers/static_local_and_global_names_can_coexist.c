// vybe-test: c/value_qualifiers/static_local_and_global_names_can_coexist
// origin: languages/c/tests/c/test_value_qualifiers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int value = 3; int sample(void) { static int value = 5; return value; }
int main() {
const char *__w[] = {"5 3\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", sample(), value);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

