// vybe-test: c/lang_bitfield_member_widths/bitfield_with_regular_int_field
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { int id; unsigned a : 1; unsigned b : 3; };
int main() {
const char *__w[] = {"9 1 4\n"};
int __n = 1, __i = 0;
struct F f; f.id = 9; f.a = 1; f.b = 4; { char __t[512]; snprintf(__t, sizeof(__t), "%d %u %u\n", f.id, f.a, f.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

