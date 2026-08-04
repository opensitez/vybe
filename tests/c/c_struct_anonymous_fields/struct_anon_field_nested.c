// vybe-test: c/c_struct_anonymous_fields/struct_anon_field_nested
// origin: languages/c/tests/c/test_c_struct_anonymous_fields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct A { int x; }; struct B { struct A; int y; }; struct C { struct B; int z; }; int main() {const char *__w[] = {"6"};
int __n = 1, __i = 0;
 struct C c; c.x = 1; c.y = 2; c.z = 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d", c.x+c.y+c.z);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

