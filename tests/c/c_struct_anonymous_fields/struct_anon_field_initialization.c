// vybe-test: c/c_struct_anonymous_fields/struct_anon_field_initialization
// origin: languages/c/tests/c/test_c_struct_anonymous_fields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { struct { int a; int b; }; int c; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct S s = { {1, 2}, 3 }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

