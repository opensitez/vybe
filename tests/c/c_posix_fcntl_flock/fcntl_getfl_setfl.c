// vybe-test: c/c_posix_fcntl_flock/fcntl_getfl_setfl
// origin: languages/c/tests/c/test_c_posix_fcntl_flock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = open("test_fcntl3.txt", O_CREAT|O_WRONLY, 0644); int flg1 = fcntl(fd, F_GETFL); fcntl(fd, F_SETFL, flg1 | O_APPEND); int flg2 = fcntl(fd, F_GETFL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (flg2 & O_APPEND) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_fcntl3.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

