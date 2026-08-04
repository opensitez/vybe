// vybe-test: c/long_types/unsigned_long_basic
// origin: languages/c/tests/c/test_long_types.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"4000000000\n"};
int __n = 1, __i = 0;
unsigned long x = 4000000000UL;
{ char __t[512]; snprintf(__t, sizeof(__t), "%lu\n", x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

