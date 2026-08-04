// vybe-test: c/c_compound_literals_basic/compound_literal_address_of
// origin: languages/c/tests/c/test_c_compound_literals_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int a; }; int main() {const char *__w[] = {"12"};
int __n = 1, __i = 0;
 struct S *p = &(struct S){12}; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p->a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

