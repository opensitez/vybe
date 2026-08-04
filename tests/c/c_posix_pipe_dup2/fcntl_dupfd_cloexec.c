// vybe-test: c/c_posix_pipe_dup2/fcntl_dupfd_cloexec
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = open("test_fcntl_dup_clo.txt", O_CREAT, 0644); int fd2 = fcntl(fd, F_DUPFD_CLOEXEC, 100); int flg = fcntl(fd2, F_GETFD); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (flg & FD_CLOEXEC) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); close(fd2); unlink("test_fcntl_dup_clo.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

