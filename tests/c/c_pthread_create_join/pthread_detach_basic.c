// vybe-test: c/c_pthread_create_join/pthread_detach_basic
// origin: languages/c/tests/c/test_c_pthread_create_join.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
#include <unistd.h>
void* f(void* a) { return NULL; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_t t; pthread_create(&t, NULL, f, NULL); int r = pthread_detach(t); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } sleep(1); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

