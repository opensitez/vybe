// vybe-test: c/c_pthread_tls_keys/getspecific_main_thread
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int main() {const char *__w[] = {"99"};
int __n = 1, __i = 0;
 pthread_key_t k; pthread_key_create(&k, NULL); pthread_setspecific(k, (void*)99); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", (long)pthread_getspecific(k));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } pthread_key_delete(k); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

