// vybe-test: c/c_enum_anonymous/enum_anonymous_in_struct
// origin: languages/c/tests/c/test_c_enum_anonymous.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { enum { A=1, B=2 } e; }; int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 struct S s; s.e = B; { char __t[512]; snprintf(__t, sizeof(__t), "%d", s.e);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

