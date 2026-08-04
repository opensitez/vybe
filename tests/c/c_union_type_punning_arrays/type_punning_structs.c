// vybe-test: c/c_union_type_punning_arrays/type_punning_structs
// origin: languages/c/tests/c/test_c_union_type_punning_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct A { int type; int val; }; struct B { int type; float f; }; union U { struct A a; struct B b; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 union U u; u.a.type = 1; u.a.val = 5; { char __t[512]; snprintf(__t, sizeof(__t), "%d", u.b.type);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

