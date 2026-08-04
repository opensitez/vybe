// vybe-test: c/c_union_anonymous/union_anon_nested_in_union
// origin: languages/c/tests/c/test_c_union_anonymous.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Outer { union { int i; char c; }; double d; }; int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 union Outer o; o.i = 42; { char __t[512]; snprintf(__t, sizeof(__t), "%d", o.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

