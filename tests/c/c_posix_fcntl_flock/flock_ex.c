// vybe-test: c/c_posix_fcntl_flock/flock_ex
// origin: languages/c/tests/c/test_c_posix_fcntl_flock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#define _DEFAULT_SOURCE
#include <sys/file.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd = open("test_flock2.txt", O_CREAT|O_RDWR, 0644); int r = flock(fd, LOCK_EX); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_flock2.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

