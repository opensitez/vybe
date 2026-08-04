// vybe-test: c/c_posix_fork_waitpid/fork_shared_file_descriptor
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
#include <fcntl.h>
int main() {const char *__w[] = {"ABC"};
int __n = 1, __i = 0;
 int fd = open("test_fork_fd.txt", O_CREAT|O_WRONLY|O_TRUNC, 0644); write(fd, "A", 1); pid_t p = fork(); if (p == 0) { write(fd, "B", 1); _exit(0); } wait(NULL); write(fd, "C", 1); close(fd); FILE *f = fopen("test_fork_fd.txt", "r"); char b[4]={0}; fread(b, 1, 3, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_fork_fd.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

