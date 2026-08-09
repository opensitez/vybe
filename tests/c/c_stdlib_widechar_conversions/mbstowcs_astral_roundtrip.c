// vybe-test: c/c_stdlib_widechar_conversions/mbstowcs_astral_roundtrip
// origin: languages/c/tests/c/test_c_stdlib_widechar_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <wchar.h>
#include <stdlib.h>
int main() {const char *__w[] = {"3 119070 3 aé\U0001D11E"};
int __n = 1, __i = 0;
 wchar_t w[16]; int n = mbstowcs(w, "aé\U0001D11E", 16); char back[32]; wcstombs(back, w, 32); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %s", n, (int)w[2], (int)wcslen(w), back);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }
