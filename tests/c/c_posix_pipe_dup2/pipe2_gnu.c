// vybe-test: c/c_posix_pipe_dup2/pipe2_gnu
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd[2]; pipe2(fd, O_CLOEXEC); int flags = fcntl(fd[0], F_GETFD); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (flags & FD_CLOEXEC) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd[0]); close(fd[1]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

