// vybe-test: c/c_stdio_vprintf_family/vdprintf_basic
// origin: languages/c/tests/c/test_c_stdio_vprintf_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdarg.h>
#include <unistd.h>
#include <fcntl.h>
void wrap_dprintf(int fd, const char *fmt, ...) { va_list args; va_start(args, fmt); vdprintf(fd, fmt, args); va_end(args); }
int main() {const char *__w[] = {"test 99"};
int __n = 1, __i = 0;
 int fd = open("test_vdprintf.txt", O_CREAT | O_WRONLY, 0644); wrap_dprintf(fd, "test %d", 99); close(fd); FILE *f = fopen("test_vdprintf.txt", "r"); char buf[20]; fgets(buf, 20, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

