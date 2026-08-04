// vybe-test: c/stdint/uint64_t_max
// origin: languages/c/tests/c/test_stdint.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"18000000000\n"};
int __n = 1, __i = 0;
uint64_t x = 18000000000ULL;
{ char __t[512]; snprintf(__t, sizeof(__t), "%llu\n", (unsigned long long)x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

