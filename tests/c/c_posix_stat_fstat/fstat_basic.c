// vybe-test: c/c_posix_stat_fstat/fstat_basic
// origin: languages/c/tests/c/test_c_posix_stat_fstat.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/stat.h>
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 int fd = open("test_fstat.txt", O_CREAT|O_WRONLY, 0644); write(fd, "xx", 2); struct stat st; int r = fstat(fd, &st); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r == 0, st.st_size == 2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_fstat.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

