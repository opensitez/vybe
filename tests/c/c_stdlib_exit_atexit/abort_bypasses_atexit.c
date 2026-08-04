// vybe-test: c/c_stdlib_exit_atexit/abort_bypasses_atexit
// origin: languages/c/tests/c/test_c_stdlib_exit_atexit.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"called"};
static int __n = 1, __i = 0;
#include <stdlib.h>
#include <signal.h>
void func() { { char __t[512]; snprintf(__t, sizeof(__t), "called");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } }
void sighandler(int sig) { exit(0); }
int main() { signal(SIGABRT, sighandler); atexit(func); abort(); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

