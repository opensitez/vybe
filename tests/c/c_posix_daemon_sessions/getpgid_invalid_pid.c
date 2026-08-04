// vybe-test: c/c_posix_daemon_sessions/getpgid_invalid_pid
// origin: languages/c/tests/c/test_c_posix_daemon_sessions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _XOPEN_SOURCE 500
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pid_t g = getpgid(-99999); { char __t[512]; snprintf(__t, sizeof(__t), "%d", g == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

