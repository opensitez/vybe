// vybe-test: c/stdint/int64_t_range
// origin: languages/c/tests/c/test_stdint.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"9000000000\n"};
int __n = 1, __i = 0;
int64_t x = 9000000000LL;
{ char __t[512]; snprintf(__t, sizeof(__t), "%lld\n", (long long)x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

