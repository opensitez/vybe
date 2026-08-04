// vybe-test: c/bit_manipulation/swap_bytes
// origin: languages/c/tests/c/test_bit_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"3412\n"};
int __n = 1, __i = 0;
unsigned short x = 0x1234;
unsigned short swapped = ((x & 0xFF) << 8) | ((x >> 8) & 0xFF);
{ char __t[512]; snprintf(__t, sizeof(__t), "%x\n", swapped);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

