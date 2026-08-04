// vybe-test: c/lang_bitfield_member_widths/bitfield_initializer_one_and_three_bits
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { unsigned a : 1; unsigned b : 3; };
int main() {
const char *__w[] = {"1 6\n"};
int __n = 1, __i = 0;
struct F f = {1, 6}; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", f.a, f.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

