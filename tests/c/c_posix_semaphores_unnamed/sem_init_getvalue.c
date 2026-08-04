// vybe-test: c/c_posix_semaphores_unnamed/sem_init_getvalue
// origin: languages/c/tests/c/test_c_posix_semaphores_unnamed.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 sem_t s; int r = sem_init(&s, 0, 42); if(r == 0) { int val = 0; sem_getvalue(&s, &val); { char __t[512]; snprintf(__t, sizeof(__t), "%d", val == 42);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } sem_destroy(&s); } else { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

