// vybe-test: c/c_posix_pipe_dup2/dup2_basic
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"d2"};
int __n = 1, __i = 0;
 int fd = open("test_dup2.txt", O_CREAT|O_WRONLY, 0644); dup2(fd, 100); write(100, "d2", 2); close(fd); close(100); FILE *f = fopen("test_dup2.txt", "r"); char buf[5]={0}; fread(buf, 1, 2, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", buf);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_dup2.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

