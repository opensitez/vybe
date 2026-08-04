// vybe-test: c/lang_bitfield_member_widths/bitfield_three_bit_and_one_bit_together
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { unsigned lo : 3; unsigned hi : 1; };
int main() {
const char *__w[] = {"4 1\n"};
int __n = 1, __i = 0;
struct F f; f.lo = 4; f.hi = 1; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", f.lo, f.hi);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

