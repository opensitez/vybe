// vybe-test: c/implicit_conversions/signed_unsigned_comparison
// origin: languages/c/tests/c/test_implicit_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"1\n"};
int __n = 1, __i = 0;
unsigned int u = 1;
int s = -1;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", u > (unsigned)s ? 0 : 1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

