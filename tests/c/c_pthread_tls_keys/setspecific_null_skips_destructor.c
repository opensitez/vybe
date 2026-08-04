// vybe-test: c/c_pthread_tls_keys/setspecific_null_skips_destructor
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int val = 0;
void d(void* a) { val = 1; }
void* f(void* a) { pthread_key_t *k = a; pthread_setspecific(*k, (void*)42); pthread_setspecific(*k, NULL); return NULL; }
int main() {const char *__w[] = {"0"};
int __n = 1, __i = 0;
 pthread_key_t k; pthread_key_create(&k, d); pthread_t t; pthread_create(&t, NULL, f, &k); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } pthread_key_delete(k); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

