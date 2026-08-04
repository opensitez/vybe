// vybe-test: c/c_posix_signal_raise/pause_basic
// origin: languages/c/tests/c/test_c_posix_signal_raise.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <unistd.h>
#include <signal.h>
#include <sys/wait.h>
void h(int s) { _exit(5); }
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 pid_t p = fork(); if(p==0) { signal(SIGUSR1, h); pause(); _exit(0); } sleep(1); kill(p, SIGUSR1); int st; waitpid(p, &st, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", WEXITSTATUS(st));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

