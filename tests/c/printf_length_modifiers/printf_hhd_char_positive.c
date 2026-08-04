// vybe-test: c/printf_length_modifiers/printf_hhd_char_positive
// origin: languages/c/tests/c/test_printf_length_modifiers.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stddef.h>
#include <inttypes.h>
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
signed char sc=42; { char __t[512]; snprintf(__t, sizeof(__t), "%hhd\n", sc);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

