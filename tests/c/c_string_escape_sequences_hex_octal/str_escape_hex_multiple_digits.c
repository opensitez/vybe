// vybe-test: c/c_string_escape_sequences_hex_octal/str_escape_hex_multiple_digits
// origin: languages/c/tests/c/test_c_string_escape_sequences_hex_octal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"A"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%c", '\x0041');
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

