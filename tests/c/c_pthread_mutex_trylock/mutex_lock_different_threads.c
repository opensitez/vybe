// vybe-test: c/c_pthread_mutex_trylock/mutex_lock_different_threads
// origin: languages/c/tests/c/test_c_pthread_mutex_trylock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
#include <unistd.h>
pthread_mutex_t m = PTHREAD_MUTEX_INITIALIZER;
int val = 0;
void* f(void* a) { pthread_mutex_lock(&m); val = 2; pthread_mutex_unlock(&m); return NULL; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 pthread_mutex_lock(&m); pthread_t t; pthread_create(&t, NULL, f, NULL); sleep(1); val = 1; pthread_mutex_unlock(&m); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

