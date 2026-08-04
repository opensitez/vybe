// vybe-test: c/c_posix_pipe_dup2/dup2_stdout_redirect
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"redirected"};
int __n = 1, __i = 0;
 int fd = open("test_redir.txt", O_CREAT|O_WRONLY, 0644); pid_t p = fork(); if (p == 0) { dup2(fd, STDOUT_FILENO); close(fd); execl("/bin/echo", "echo", "redirected", NULL); _exit(1); } wait(NULL); close(fd); FILE *f = fopen("test_redir.txt", "r"); char buf[20]={0}; fread(buf, 1, 10, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_redir.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

