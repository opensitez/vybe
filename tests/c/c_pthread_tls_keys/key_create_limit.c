// vybe-test: c/c_pthread_tls_keys/key_create_limit
// origin: languages/c/tests/c/test_c_pthread_tls_keys.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
#include <limits.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 /* PTHREAD_KEYS_MAX is often 1024 or more, test we can create at least 10 */ pthread_key_t keys[10]; int ok = 1; for(int i=0; i<10; i++) if(pthread_key_create(&keys[i], NULL) != 0) ok = 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d", ok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } for(int i=0; i<10; i++) pthread_key_delete(keys[i]); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

