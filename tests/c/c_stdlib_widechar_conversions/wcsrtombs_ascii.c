// vybe-test: c/c_stdlib_widechar_conversions/wcsrtombs_ascii
// origin: languages/c/tests/c/test_c_stdlib_widechar_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <wchar.h>
int main() {const char *__w[] = {"2 1 h"};
int __n = 1, __i = 0;
 char s[10]; const wchar_t *src = L"hi"; mbstate_t st = {0}; size_t len = wcsrtombs(s, &src, 10, &st); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %c", (int)len, src == NULL, s[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

