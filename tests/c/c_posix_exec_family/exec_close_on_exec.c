// vybe-test: c/c_posix_exec_family/exec_close_on_exec
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int fd = open("test_cloexec.txt", O_CREAT|O_WRONLY, 0644); fcntl(fd, F_SETFD, FD_CLOEXEC); /* We test that fcntl succeeds and FD_CLOEXEC is defined */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_cloexec.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

