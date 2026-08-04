// vybe-test: c/c_posix_pipe_dup2/dup_basic
// origin: languages/c/tests/c/test_c_posix_pipe_dup2.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"dup 1"};
int __n = 1, __i = 0;
 int fd = open("test_dup.txt", O_CREAT|O_WRONLY, 0644); int fd2 = dup(fd); write(fd2, "dup", 3); close(fd); close(fd2); FILE *f = fopen("test_dup.txt", "r"); char buf[5]={0}; fread(buf, 1, 3, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s %d", buf, fd != fd2);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_dup.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

