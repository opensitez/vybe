// vybe-test: c/c_stdio_vprintf_family/vsnprintf_zero_size
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
int wrap(char *buf, size_t n, const char *fmt, ...) { va_list args; va_start(args, fmt); int res = vsnprintf(buf, n, fmt, args); va_end(args); return res; }
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 int len = wrap(NULL, 0, "12345"); { char __t[512]; snprintf(__t, sizeof(__t), "%d", len);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

