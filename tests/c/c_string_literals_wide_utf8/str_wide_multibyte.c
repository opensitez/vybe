// vybe-test: c/c_string_literals_wide_utf8/str_wide_multibyte
// origin: languages/c/tests/c/test_c_string_literals_wide_utf8.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stddef.h>
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 wchar_t str[] = L"\u00A9"; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)(sizeof(str)/sizeof(wchar_t)));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

