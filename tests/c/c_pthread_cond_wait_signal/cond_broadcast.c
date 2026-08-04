// vybe-test: c/c_pthread_cond_wait_signal/cond_broadcast
// origin: languages/c/tests/c/test_c_pthread_cond_wait_signal.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
#include <unistd.h>
pthread_mutex_t m = PTHREAD_MUTEX_INITIALIZER;
pthread_cond_t c = PTHREAD_COND_INITIALIZER;
int ready = 0, count = 0;
void* f(void* a) { pthread_mutex_lock(&m); while(!ready) pthread_cond_wait(&c, &m); count++; pthread_mutex_unlock(&m); return NULL; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); sleep(1); pthread_mutex_lock(&m); ready = 1; pthread_cond_broadcast(&c); pthread_mutex_unlock(&m); pthread_join(t1, NULL); pthread_join(t2, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

