// vybe-test: c/wchar_string_operations/wcstombs_narrow_first_byte
// origin: languages/c/tests/c/test_wchar_string_operations.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <wchar.h>
#include <stdlib.h>
int main() {
const char *__w[] = {"Z\n"};
int __n = 1, __i = 0;
char b[8]; wcstombs(b, L"Z", 8); { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", b[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

