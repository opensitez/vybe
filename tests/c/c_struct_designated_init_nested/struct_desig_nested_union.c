// vybe-test: c/c_struct_designated_init_nested/struct_desig_nested_union
// origin: languages/c/tests/c/test_c_struct_designated_init_nested.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union U { int i; char c; }; struct S { union U u; }; int main() {const char *__w[] = {"65"};
int __n = 1, __i = 0;
 struct S s = { .u.i = 65 }; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.u.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

