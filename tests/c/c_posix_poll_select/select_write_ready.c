// vybe-test: c/c_posix_poll_select/select_write_ready
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/select.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 int fd[2]; pipe(fd); fd_set w; FD_ZERO(&w); FD_SET(fd[1], &w); struct timeval tv = {0, 0}; int res = select(fd[1]+1, NULL, &w, NULL, &tv); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", res == 1, FD_ISSET(fd[1], &w) != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd[0]); close(fd[1]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

