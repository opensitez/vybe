// vybe-test: c/c_pthread_create_join/pthread_cleanup_push_pop
// origin: languages/c/tests/c/test_c_pthread_create_join.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int val = 0;
void c(void* a) { val = 1; }
void* f(void* a) { pthread_cleanup_push(c, NULL); pthread_cleanup_pop(1); return NULL; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_t t; pthread_create(&t, NULL, f, NULL); pthread_join(t, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

