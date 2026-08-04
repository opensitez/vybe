// vybe-test: c/c_struct_bitfields_signedness/bitfield_sign_plain_int_impl_defined
// origin: languages/c/tests/c/test_c_struct_bitfields_signedness.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int a:3; }; int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 struct S s; s.a = 7; int val = s.a; { char __t[512]; snprintf(__t, sizeof(__t), "%d", val == 7 || val == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

