// vybe-test: c/c_pthread_create_join/pthread_setcanceltype
// origin: languages/c/tests/c/test_c_pthread_create_join.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
void* f(void* a) { int old; pthread_setcanceltype(PTHREAD_CANCEL_ASYNCHRONOUS, &old); return (void*)(long)old; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_t t; pthread_create(&t, NULL, f, NULL); void *r; pthread_join(t, &r); { char __t[512]; snprintf(__t, sizeof(__t), "%d", (long)r == PTHREAD_CANCEL_DEFERRED || (long)r == PTHREAD_CANCEL_ASYNCHRONOUS);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

