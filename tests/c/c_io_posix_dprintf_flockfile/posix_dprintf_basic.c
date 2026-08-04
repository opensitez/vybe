// vybe-test: c/c_io_posix_dprintf_flockfile/posix_dprintf_basic
// origin: languages/c/tests/c/test_c_io_posix_dprintf_flockfile.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <stdio.h>
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"hello 123"};
int __n = 1, __i = 0;
 int fd = open("test_dprintf.txt", O_CREAT|O_WRONLY, 0644); if (fd != -1) { dprintf(fd, "hello %d", 123); close(fd); FILE *f = fopen("test_dprintf.txt", "r"); char buf[20]; fgets(buf, sizeof(buf), f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

