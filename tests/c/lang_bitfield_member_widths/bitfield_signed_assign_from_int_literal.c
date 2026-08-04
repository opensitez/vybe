// vybe-test: c/lang_bitfield_member_widths/bitfield_signed_assign_from_int_literal
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { signed s : 4; };
int main() {
const char *__w[] = {"-8\n"};
int __n = 1, __i = 0;
struct F f; f.s = -8; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", f.s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

