// vybe-test: c/c_posix_fork_waitpid/wait3_wait4_compile
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#define _DEFAULT_SOURCE
#include <sys/wait.h>
#include <sys/resource.h>
#include <unistd.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 pid_t p = fork(); if (p == 0) _exit(0); int st; struct rusage ru; wait4(p, &st, 0, &ru); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

