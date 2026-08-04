// vybe-test: c/c_pthread_barrier/barrier_two_phases
// origin: languages/c/tests/c/test_c_pthread_barrier.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <unistd.h>
pthread_barrier_t b;
int p = 0;
void* f(void* a) { pthread_barrier_wait(&b); p++; pthread_barrier_wait(&b); return NULL; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_barrier_init(&b, NULL, 2); pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_barrier_wait(&b); pthread_barrier_wait(&b); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

