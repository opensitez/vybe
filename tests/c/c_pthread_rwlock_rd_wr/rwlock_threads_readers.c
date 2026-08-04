// vybe-test: c/c_pthread_rwlock_rd_wr/rwlock_threads_readers
// origin: languages/c/tests/c/test_c_pthread_rwlock_rd_wr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <unistd.h>
pthread_rwlock_t rw = PTHREAD_RWLOCK_INITIALIZER;
void* f(void* a) { pthread_rwlock_rdlock(&rw); sleep(1); pthread_rwlock_unlock(&rw); return NULL; }
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); pthread_join(t1, NULL); pthread_join(t2, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

