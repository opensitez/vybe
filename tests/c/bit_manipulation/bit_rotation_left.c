// vybe-test: c/bit_manipulation/bit_rotation_left
// origin: languages/c/tests/c/test_bit_manipulation.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"23456781\n"};
int __n = 1, __i = 0;
unsigned int x = 0x12345678;
unsigned int rot = (x << 4) | (x >> 28);
{ char __t[512]; snprintf(__t, sizeof(__t), "%x\n", rot);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

