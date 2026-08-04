// vybe-test: c/wchar_string_operations/wcscpy_first_char_narrow
// origin: languages/c/tests/c/test_wchar_string_operations.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
int main() {
const char *__w[] = {"h\n"};
int __n = 1, __i = 0;
wchar_t d[4]; wcscpy(d, L"hi"); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", (char)wctob(d[0]));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

