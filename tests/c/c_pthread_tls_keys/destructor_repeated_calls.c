// vybe-test: c/c_pthread_tls_keys/destructor_repeated_calls
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int iters = 0;
pthread_key_t k;
void d(void* a) { iters++; if(iters < 3) pthread_setspecific(k, (void*)1); }
void* f(void* a) { pthread_setspecific(k, (void*)1); return NULL; }
int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 pthread_key_create(&k, d); pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", iters);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } pthread_key_delete(k); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

