// vybe-test: c/c_posix_poll_select/fd_set_macros
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/select.h>
int main() {const char *__w[] = {"1 0"};
int __n = 1, __i = 0;
 fd_set s; FD_ZERO(&s); FD_SET(5, &s); int r1 = FD_ISSET(5, &s); FD_CLR(5, &s); int r2 = FD_ISSET(5, &s); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r1 != 0, r2 != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

