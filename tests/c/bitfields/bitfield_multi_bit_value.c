// vybe-test: c/bitfields/bitfield_multi_bit_value
// origin: languages/c/tests/c/test_bitfields.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Packed { unsigned int low : 4; unsigned int high : 4; };
int main() {
const char *__w[] = {"5 12\n"};
int __n = 1, __i = 0;
struct Packed p;
p.low = 5; p.high = 12;
{ char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", p.low, p.high);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

