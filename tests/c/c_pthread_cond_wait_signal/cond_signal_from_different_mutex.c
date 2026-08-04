// vybe-test: c/c_pthread_cond_wait_signal/cond_signal_from_different_mutex
// origin: languages/c/tests/c/test_c_pthread_cond_wait_signal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* UB to wait on same cond with different mutexes concurrently, but signaling from anywhere is ok */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

