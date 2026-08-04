// vybe-test: c/c_pthread_tls_keys/multiple_keys
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int main() {const char *__w[] = {"1 2"};
int __n = 1, __i = 0;
 pthread_key_t k1, k2; pthread_key_create(&k1, NULL); pthread_key_create(&k2, NULL); pthread_setspecific(k1, (void*)1); pthread_setspecific(k2, (void*)2); { char __t[512]; snprintf(__t, sizeof(__t), "%ld %ld", (long)pthread_getspecific(k1), (long)pthread_getspecific(k2));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } pthread_key_delete(k1); pthread_key_delete(k2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

