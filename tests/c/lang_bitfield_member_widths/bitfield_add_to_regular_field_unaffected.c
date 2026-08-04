// vybe-test: c/lang_bitfield_member_widths/bitfield_add_to_regular_field_unaffected
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { int n; unsigned a : 3; };
int main() {
const char *__w[] = {"15 3\n"};
int __n = 1, __i = 0;
struct F f; f.n = 10; f.a = 3; f.n += 5; { char __t[512]; snprintf(__t, sizeof(__t), "%d %u\n", f.n, f.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

