// vybe-test: c/c_stdio_vprintf_family/vsscanf_basic
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdarg.h>
void wrap_sscanf(const char *str, const char *fmt, ...) { va_list args; va_start(args, fmt); vsscanf(str, fmt, args); va_end(args); }
int main() {const char *__w[] = {"10 20"};
int __n = 1, __i = 0;
 int val1, val2; wrap_sscanf("10 20", "%d %d", &val1, &val2); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", val1, val2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

