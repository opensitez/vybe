// vybe-test: c/c_pthread_barrier/barrier_wait_returns_serial_once
// origin: languages/c/tests/c/test_c_pthread_barrier.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
pthread_barrier_t b;
void* f(void* a) { return (void*)(long)pthread_barrier_wait(&b); }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_barrier_init(&b, NULL, 2); pthread_t t; pthread_create(&t, NULL, f, NULL); int r1 = pthread_barrier_wait(&b); void *res; pthread_join(t, &res); int r2 = (long)res; { char __t[512]; snprintf(__t, sizeof(__t), "%d", (r1 == PTHREAD_BARRIER_SERIAL_THREAD && r2 != PTHREAD_BARRIER_SERIAL_THREAD) || (r2 == PTHREAD_BARRIER_SERIAL_THREAD && r1 != PTHREAD_BARRIER_SERIAL_THREAD));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

