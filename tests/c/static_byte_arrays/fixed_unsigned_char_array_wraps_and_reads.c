// vybe-test: c/static_byte_arrays/fixed_unsigned_char_array_wraps_and_reads
// origin: regression for fixed unsigned-byte array lowering
#include <stdio.h>
#include <string.h>
#include <assert.h>

static unsigned char screen[320 * 200];

int main() {
const char *__w[] = {"44 255\n"};
int __n = 1, __i = 0;
screen[0] = (unsigned char)300;
screen[319 + 199 * 320] = (unsigned char)255;
{ char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", screen[0], screen[319 + 199 * 320]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
