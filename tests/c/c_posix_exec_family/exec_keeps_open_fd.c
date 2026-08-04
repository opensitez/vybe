// vybe-test: c/c_posix_exec_family/exec_keeps_open_fd
// origin: languages/c/tests/c/test_c_posix_exec_family.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <fcntl.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"hi"};
int __n = 1, __i = 0;
 int fd = open("test_keep_fd.txt", O_CREAT|O_WRONLY, 0644); pid_t p = fork(); if(p==0) { /* fd remains open across exec unless cloexec is set */ char buf[50]; sprintf(buf, "echo hi >&%d", fd); execl("/bin/sh", "sh", "-c", buf, NULL); _exit(1); } waitpid(p, NULL, 0); close(fd); FILE *f = fopen("test_keep_fd.txt", "r"); char b[10]={0}; fread(b, 1, 9, f); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } fclose(f); unlink("test_keep_fd.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

