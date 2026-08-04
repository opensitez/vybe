// vybe-test: c/c_stdio_vprintf_family/vsscanf_eof
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
int wrap(const char *str, const char *fmt, ...) { va_list args; va_start(args, fmt); int res = vsscanf(str, fmt, args); va_end(args); return res; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int v; int n = wrap("", "%d", &v); { char __t[512]; snprintf(__t, sizeof(__t), "%d", n == EOF);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

