// vybe-test: c/c_pthread_mutex_trylock/mutexattr_gettype
// origin: languages/c/tests/c/test_c_pthread_mutex_trylock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <pthread.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 pthread_mutexattr_t a; pthread_mutexattr_init(&a); pthread_mutexattr_settype(&a, PTHREAD_MUTEX_RECURSIVE); int type; pthread_mutexattr_gettype(&a, &type); { char __t[512]; snprintf(__t, sizeof(__t), "%d", type == PTHREAD_MUTEX_RECURSIVE);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

