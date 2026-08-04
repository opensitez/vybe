// vybe-test: c/c_posix_poll_select/poll_zero_nfds
// origin: languages/c/tests/c/test_c_posix_poll_select.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <poll.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int res = poll(NULL, 0, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

