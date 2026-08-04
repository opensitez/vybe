// vybe-test: c/typedefs/typedef_alias_can_shadow_builtin_type_name_locally
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Count;
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
Count count = 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

