// vybe-test: c/c_struct_anonymous_fields/struct_anon_union_in_union
// origin: languages/c/tests/c/test_c_struct_anonymous_fields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union U { union { int a; int b; }; float c; }; int main() {const char *__w[] = {"10"};
int __n = 1, __i = 0;
 union U u; u.a = 10; { char __t[512]; snprintf(__t, sizeof(__t), "%d", u.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

