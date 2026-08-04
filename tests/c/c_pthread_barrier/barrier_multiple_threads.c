// vybe-test: c/c_pthread_barrier/barrier_multiple_threads
// origin: languages/c/tests/c/test_c_pthread_barrier.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <unistd.h>
pthread_barrier_t b;
int count = 0;
pthread_mutex_t m = PTHREAD_MUTEX_INITIALIZER;
void* f(void* a) { pthread_barrier_wait(&b); pthread_mutex_lock(&m); count++; pthread_mutex_unlock(&m); return NULL; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 pthread_barrier_init(&b, NULL, 3); pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); pthread_barrier_wait(&b); pthread_join(t1, NULL); pthread_join(t2, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

