// vybe-test: c/c_pthread_spinlock/spin_trylock_stress
// origin: languages/c/tests/c/test_c_pthread_spinlock.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <pthread.h>
pthread_spinlock_t s;
int count = 0;
void* f(void* a) { for(int i=0; i<1000; i++) { while(pthread_spin_trylock(&s) != 0); count++; pthread_spin_unlock(&s); } return NULL; }
int main() {const char *__w[] = {"2000"};
int __n = 1, __i = 0;
 pthread_spin_init(&s, PTHREAD_PROCESS_PRIVATE); pthread_t t1, t2; pthread_create(&t1, NULL, f, NULL); pthread_create(&t2, NULL, f, NULL); pthread_join(t1, NULL); pthread_join(t2, NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

