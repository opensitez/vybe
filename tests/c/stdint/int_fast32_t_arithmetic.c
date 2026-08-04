// vybe-test: c/stdint/int_fast32_t_arithmetic
// origin: languages/c/tests/c/test_stdint.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"300\n"};
int __n = 1, __i = 0;
int_fast32_t a = 100;
int_fast32_t b = 200;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)(a + b));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

