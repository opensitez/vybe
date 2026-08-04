// vybe-test: c/c_stdio_vprintf_family/vasprintf_basic
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <stdarg.h>
#include <stdlib.h>
void wrap_asprintf(char **strp, const char *fmt, ...) { va_list args; va_start(args, fmt); vasprintf(strp, fmt, args); va_end(args); }
int main() {const char *__w[] = {"hello 123"};
int __n = 1, __i = 0;
 char *str; wrap_asprintf(&str, "hello %d", 123); { char __t[512]; snprintf(__t, sizeof(__t), "%s", str);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(str); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

