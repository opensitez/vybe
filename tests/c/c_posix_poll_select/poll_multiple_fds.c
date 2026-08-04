// vybe-test: c/c_posix_poll_select/poll_multiple_fds
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <poll.h>
#include <unistd.h>
int main() {const char *__w[] = {"1 1 1"};
int __n = 1, __i = 0;
 int f1[2], f2[2]; pipe(f1); pipe(f2); write(f1[1], "x", 1); struct pollfd p[2] = {{f1[0], POLLIN, 0}, {f2[0], POLLIN, 0}}; int res = poll(p, 2, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", res == 1, (p[0].revents & POLLIN) != 0, (p[1].revents & POLLIN) == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(f1[0]); close(f1[1]); close(f2[0]); close(f2[1]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

