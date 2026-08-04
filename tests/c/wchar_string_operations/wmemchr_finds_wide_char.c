// vybe-test: c/wchar_string_operations/wmemchr_finds_wide_char
// origin: languages/c/tests/c/test_wchar_string_operations.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
int main() {
const char *__w[] = {"b\n"};
int __n = 1, __i = 0;
wchar_t *p = wmemchr(L"abc", L'b', 3); { char __t[512]; snprintf(__t, sizeof(__t), "%lc\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

