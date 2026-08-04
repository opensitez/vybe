// vybe-test: c/long_types/long_long_multiplication
// origin: languages/c/tests/c/test_long_types.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"10000000000\n"};
int __n = 1, __i = 0;
long long a = 100000LL;
long long b = 100000LL;
{ char __t[512]; snprintf(__t, sizeof(__t), "%lld\n", a * b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

