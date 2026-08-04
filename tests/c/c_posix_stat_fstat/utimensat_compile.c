// vybe-test: c/c_posix_stat_fstat/utimensat_compile
// origin: languages/c/tests/c/test_c_posix_stat_fstat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 /* AT_FDCWD */ int fd = open("test_utime.txt", O_CREAT|O_WRONLY, 0644); close(fd); struct timespec ts[2] = {{0, UTIME_NOW}, {0, UTIME_NOW}}; int r = utimensat(AT_FDCWD, "test_utime.txt", ts, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } unlink("test_utime.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

