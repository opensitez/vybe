// vybe-test: c/lang_bitfield_member_widths/bitfield_in_struct_array
// origin: languages/c/tests/c/test_lang_bitfield_member_widths.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct F { unsigned a : 3; };
int main() {
const char *__w[] = {"1 7\n"};
int __n = 1, __i = 0;
struct F arr[2]; arr[0].a = 1; arr[1].a = 7; { char __t[512]; snprintf(__t, sizeof(__t), "%u %u\n", arr[0].a, arr[1].a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

