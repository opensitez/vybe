// vybe-test: c/bit_manipulation/count_set_bits
// origin: languages/c/tests/c/test_bit_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;

int n = 0b10110110;
int count = 0;
while (n) { count += n & 1; n >>= 1; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

