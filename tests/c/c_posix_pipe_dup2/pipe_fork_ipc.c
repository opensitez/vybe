// vybe-test: c/c_posix_pipe_dup2/pipe_fork_ipc
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"msg"};
int __n = 1, __i = 0;
 int fd[2]; pipe(fd); pid_t p = fork(); if (p == 0) { close(fd[0]); write(fd[1], "msg", 3); close(fd[1]); _exit(0); } close(fd[1]); char buf[5] = {0}; read(fd[0], buf, 3); close(fd[0]); wait(NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

