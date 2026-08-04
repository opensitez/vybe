// vybe-test: c/c_stdio_vprintf_family/vsnprintf_exact_size
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
void wrap(char *buf, size_t n, const char *fmt, ...) { va_list args; va_start(args, fmt); vsnprintf(buf, n, fmt, args); va_end(args); }
int main() {const char *__w[] = {"123"};
int __n = 1, __i = 0;
 char buf[4]; wrap(buf, 4, "123"); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

