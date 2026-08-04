// vybe-test: c/c_posix_pipe_dup2/dup2_closes_newfd
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd1 = open("test_d2_1.txt", O_CREAT, 0644); int fd2 = open("test_d2_2.txt", O_CREAT, 0644); dup2(fd1, fd2); /* fd2 was closed and reopened as fd1 duplicate */ int r = write(fd2, "x", 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == -1 || r == 1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd1); close(fd2); unlink("test_d2_1.txt"); unlink("test_d2_2.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

