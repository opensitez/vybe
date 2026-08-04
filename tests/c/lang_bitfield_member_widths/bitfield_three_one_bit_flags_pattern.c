// vybe-test: c/lang_bitfield_member_widths/bitfield_three_one_bit_flags_pattern
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { unsigned a : 1; unsigned b : 1; unsigned c : 1; };
int main() {
const char *__w[] = {"1 0 1\n"};
int __n = 1, __i = 0;
struct F f = {1, 0, 1}; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u %u\n", f.a, f.b, f.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

