// vybe-test: c/c_struct_bitfields_packing/bitfield_packing_struct_assignment
// origin: languages/c/tests/c/test_c_struct_bitfields_packing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { unsigned int a:4; unsigned int b:4; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct S s1 = {1, 2}; struct S s2 = s1; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s2.a+s2.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

