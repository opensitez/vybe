// vybe-test: c/c_pthread_spinlock/spin_destroy_locked
// origin: languages/c/tests/c/test_c_pthread_spinlock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* UB to destroy locked spinlock, compile check only */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

