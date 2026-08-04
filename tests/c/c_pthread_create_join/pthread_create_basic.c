// vybe-test: c/c_pthread_create_join/pthread_create_basic
// origin: languages/c/tests/c/test_c_pthread_create_join.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
void* f(void* a) { return a; }
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 pthread_t t; int r = pthread_create(&t, NULL, f, (void*)42); if(r==0) { void *res; pthread_join(t, &res); { char __t[512]; snprintf(__t, sizeof(__t), "%ld", (long)res);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

