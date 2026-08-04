// vybe-test: c/c_struct_bitfields_packing/bitfield_packing_pointer_to_struct
// origin: languages/c/tests/c/test_c_struct_bitfields_packing.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { unsigned int a:4; }; int main() {const char *__w[] = {"7"};
int __n = 1, __i = 0;
 struct S s = {7}; struct S *p = &s; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p->a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

