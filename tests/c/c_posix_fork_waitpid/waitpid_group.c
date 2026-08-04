// vybe-test: c/c_posix_fork_waitpid/waitpid_group
// origin: languages/c/tests/c/test_c_posix_fork_waitpid.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pid_t p = fork(); if (p == 0) _exit(0); int st; pid_t r = waitpid(0, &st, 0); /* 0 means any child in same process group */ { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

