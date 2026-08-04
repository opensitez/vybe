// vybe-test: c/typedefs/typedef_alias_for_int_can_declare_variable
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Number;
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
Number value = 7; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", value);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

