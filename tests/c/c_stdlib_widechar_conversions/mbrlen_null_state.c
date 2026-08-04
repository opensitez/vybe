// vybe-test: c/c_stdlib_widechar_conversions/mbrlen_null_state
// origin: languages/c/tests/c/test_c_stdlib_widechar_conversions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <wchar.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 size_t len = mbrlen("A", 1, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (int)len);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

