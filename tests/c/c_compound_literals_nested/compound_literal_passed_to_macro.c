// vybe-test: c/c_compound_literals_nested/compound_literal_passed_to_macro
// origin: languages/c/tests/c/test_c_compound_literals_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define GET_A(s) ((s).a)
struct S { int a; }; int main() {const char *__w[] = {"99"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", GET_A((struct S){99}));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

