// vybe-test: c/c_pthread_rwlock_rd_wr/rwlockattr_init_destroy
// origin: languages/c/tests/c/test_c_pthread_rwlock_rd_wr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
int main() {const char *__w[] = {"1 1"};
int __n = 1, __i = 0;
 pthread_rwlockattr_t a; int r1 = pthread_rwlockattr_init(&a); int r2 = pthread_rwlockattr_destroy(&a); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", r1 == 0, r2 == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

