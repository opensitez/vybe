// vybe-test: c/c_posix_fork_waitpid/waitpid_nohang
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pid_t p = fork(); if (p == 0) { sleep(2); _exit(0); } int st; int res = waitpid(p, &st, WNOHANG); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } waitpid(p, &st, 0); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

