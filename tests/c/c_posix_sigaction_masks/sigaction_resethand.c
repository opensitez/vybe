// vybe-test: c/c_posix_sigaction_masks/sigaction_resethand
// origin: languages/c/tests/c/test_c_posix_sigaction_masks.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <signal.h>
#include <stdlib.h>
int count = 0;
void h(int s) { count++; if(count==2) exit(0); }
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 struct sigaction sa; sa.sa_handler = h; sigemptyset(&sa.sa_mask); sa.sa_flags = SA_RESETHAND; sigaction(SIGUSR1, &sa, NULL); raise(SIGUSR1); /* Second raise will terminate process because handler was reset to DFL */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

