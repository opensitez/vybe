// vybe-test: c/inttypes_format_macro_output/priomax_octal_twenty
// origin: languages/c/tests/c/test_inttypes_format_macro_output.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"20\n"};
int __n = 1, __i = 0;
uintmax_t v=16; { char __t[512]; snprintf(__t, sizeof(__t), "%" PRIoMAX "\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

