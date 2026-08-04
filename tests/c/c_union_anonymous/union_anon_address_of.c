// vybe-test: c/c_union_anonymous/union_anon_address_of
// origin: languages/c/tests/c/test_c_union_anonymous.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { union { int i; float f; }; }; int main() {const char *__w[] = {"99"};
int __n = 1, __i = 0;
 struct S s; int *p = &s.i; *p = 99; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

