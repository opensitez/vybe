// vybe-test: c/inttypes_format_macro_output/pridmax_negative_one
// origin: languages/c/tests/c/test_inttypes_format_macro_output.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"-1\n"};
int __n = 1, __i = 0;
intmax_t v=-1; { char __t[512]; snprintf(__t, sizeof(__t), "%" PRIdMAX "\n", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

