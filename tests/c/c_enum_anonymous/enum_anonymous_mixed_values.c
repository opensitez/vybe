// vybe-test: c/c_enum_anonymous/enum_anonymous_mixed_values
// origin: languages/c/tests/c/test_c_enum_anonymous.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
enum { A = 5, B, C = 10, D }; int main() {const char *__w[] = {"17"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", B+D);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

