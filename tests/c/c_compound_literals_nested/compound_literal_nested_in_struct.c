// vybe-test: c/c_compound_literals_nested/compound_literal_nested_in_struct
// origin: languages/c/tests/c/test_c_compound_literals_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Inner { int x; }; struct Outer { struct Inner i; }; int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 struct Outer o = { (struct Inner){5} }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", o.i.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

