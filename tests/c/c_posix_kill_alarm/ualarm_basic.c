// vybe-test: c/c_posix_kill_alarm/ualarm_basic
// origin: languages/c/tests/c/test_c_posix_kill_alarm.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"ualarm"};
static int __n = 1, __i = 0;
#define _XOPEN_SOURCE 500
#include <unistd.h>
#include <signal.h>
void h(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "ualarm");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } _exit(0); }
int main() { signal(SIGALRM, h); ualarm(50000, 0); pause(); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

