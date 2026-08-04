// vybe-test: c/c_posix_pipe_dup2/fcntl_dupfd
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = open("test_fcntl_dup.txt", O_CREAT, 0644); int fd2 = fcntl(fd, F_DUPFD, 100); { char __t[512]; snprintf(__t, sizeof(__t), "%d", fd2 >= 100);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); close(fd2); unlink("test_fcntl_dup.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

