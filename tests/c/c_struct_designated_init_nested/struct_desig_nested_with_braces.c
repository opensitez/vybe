// vybe-test: c/c_struct_designated_init_nested/struct_desig_nested_with_braces
// origin: languages/c/tests/c/test_c_struct_designated_init_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int a; }; struct Outer { struct Inner i; }; int main() {const char *__w[] = {"7"};
int __n = 1, __i = 0;
 struct Outer o = { .i = { .a = 7 } }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", o.i.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

