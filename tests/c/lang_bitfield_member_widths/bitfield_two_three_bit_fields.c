// vybe-test: c/lang_bitfield_member_widths/bitfield_two_three_bit_fields
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { unsigned x : 3; unsigned y : 3; };
int main() {
const char *__w[] = {"3 6\n"};
int __n = 1, __i = 0;
struct F f; f.x = 3; f.y = 6; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", f.x, f.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

