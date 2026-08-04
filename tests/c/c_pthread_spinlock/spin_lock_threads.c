// vybe-test: c/c_pthread_spinlock/spin_lock_threads
// origin: languages/c/tests/c/test_c_pthread_spinlock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <unistd.h>
pthread_spinlock_t s;
int val = 0;
void* f(void* a) { pthread_spin_lock(&s); val = 2; pthread_spin_unlock(&s); return NULL; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 pthread_spin_init(&s, PTHREAD_PROCESS_PRIVATE); pthread_spin_lock(&s); pthread_t t; pthread_create(&t, NULL, f, NULL); sleep(1); val = 1; pthread_spin_unlock(&s); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

