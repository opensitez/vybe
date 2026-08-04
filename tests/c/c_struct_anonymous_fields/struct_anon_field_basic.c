// vybe-test: c/c_struct_anonymous_fields/struct_anon_field_basic
// origin: languages/c/tests/c/test_c_struct_anonymous_fields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int a; }; struct Outer { struct Inner; int b; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct Outer o; o.a = 1; o.b = 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", o.a + o.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

