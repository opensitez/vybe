// vybe-test: c/typedefs/typedef_alias_can_be_used_in_function_signature
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Number; Number double_it(Number x) { return x * 2; }
int main() {
const char *__w[] = {"12\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", double_it(6));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

