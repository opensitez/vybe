// vybe-test: c/c_stdlib_multibyte_conversions/mblen_invalid
// origin: languages/c/tests/c/test_c_stdlib_multibyte_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
#include <locale.h>
int main() {const char *__w[] = {"-1"};
int __n = 1, __i = 0;
 if (!setlocale(LC_CTYPE, "C.UTF-8")) setlocale(LC_CTYPE, "en_US.UTF-8");
 char inv[] = {(char)0xff}; int len = mblen(inv, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", len);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

