// vybe-test: c/c_posix_fcntl_flock/flock_invalid_fd
// origin: languages/c/tests/c/test_c_posix_fcntl_flock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#define _DEFAULT_SOURCE
#include <sys/file.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int r = flock(-1, LOCK_SH); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

