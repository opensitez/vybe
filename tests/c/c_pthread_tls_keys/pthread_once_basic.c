// vybe-test: c/c_pthread_tls_keys/pthread_once_basic
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
pthread_once_t once = PTHREAD_ONCE_INIT;
int count = 0;
void init() { count++; }
void* f(void* a) { pthread_once(&once, init); return NULL; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); pthread_join(t1, NULL); pthread_join(t2, NULL); pthread_once(&once, init); { char __t[512]; snprintf(__t, sizeof(__t), "%d", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

