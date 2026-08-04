// vybe-test: c/c_posix_kill_alarm/setitimer_virtual
// origin: languages/c/tests/c/test_c_posix_kill_alarm.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"vtimer"};
static int __n = 1, __i = 0;
#define _POSIX_C_SOURCE 200809L
#include <sys/time.h>
#include <signal.h>
void h(int s) { { char __t[512]; snprintf(__t, sizeof(__t), "vtimer");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } _exit(0); }
int main() { signal(SIGVTALRM, h); struct itimerval it = {0}; it.it_value.tv_usec = 50000; setitimer(ITIMER_VIRTUAL, &it, NULL); while(1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

