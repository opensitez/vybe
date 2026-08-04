// vybe-test: c/c_posix_fcntl_flock/fcntl_getlk
// origin: languages/c/tests/c/test_c_posix_fcntl_flock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <fcntl.h>
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = open("test_fcntl7.txt", O_CREAT|O_RDWR, 0644); struct flock fl = {0}; fl.l_type = F_WRLCK; fl.l_whence = SEEK_SET; fcntl(fd, F_SETLK, &fl); pid_t p = fork(); if(p==0) { struct flock fl2 = {0}; fl2.l_type = F_WRLCK; fl2.l_whence = SEEK_SET; fcntl(fd, F_GETLK, &fl2); { char __t[512]; snprintf(__t, sizeof(__t), "%d", fl2.l_type != F_UNLCK && fl2.l_pid == getppid());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } _exit(0); } wait(NULL); close(fd); unlink("test_fcntl7.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

