// vybe-test: c/c_pthread_tls_keys/tls_macro_threads_isolated
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
#include <unistd.h>
__thread int val = 0;
void* f(void* a) { val = 1; return NULL; }
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 val = 2; pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

