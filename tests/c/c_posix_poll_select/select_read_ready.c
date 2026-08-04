// vybe-test: c/c_posix_poll_select/select_read_ready
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/select.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 int fd[2]; pipe(fd); write(fd[1], "x", 1); fd_set r; FD_ZERO(&r); FD_SET(fd[0], &r); struct timeval tv = {0, 0}; int res = select(fd[0]+1, &r, NULL, NULL, &tv); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", res == 1, FD_ISSET(fd[0], &r) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd[0]); close(fd[1]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

