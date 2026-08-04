// vybe-test: c/scope_storage/inner_scope_can_declare_same_typedef_name_shadow_independently
// origin: languages/c/tests/c/test_scope_storage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int Number;
int main() {
const char *__w[] = {"4\n", "3\n"};
int __n = 2, __i = 0;
Number a = 3; { typedef int Number; Number b = 4; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

