// vybe-test: c/c_union_type_punning_arrays/type_punning_union_initialization
// origin: languages/c/tests/c/test_c_union_type_punning_arrays.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union U { int i; char c; }; int main() {const char *__w[] = {"A"};
int __n = 1, __i = 0;
 union U u = {65}; { char __t[512]; snprintf(__t, sizeof(__t), "%c", u.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

