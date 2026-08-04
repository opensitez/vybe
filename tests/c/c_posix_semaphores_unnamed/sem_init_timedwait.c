// vybe-test: c/c_posix_semaphores_unnamed/sem_init_timedwait
// origin: languages/c/tests/c/test_c_posix_semaphores_unnamed.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <time.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 sem_t s; int r = sem_init(&s, 0, 0); if(r == 0) { struct timespec ts; clock_gettime(CLOCK_REALTIME, &ts); ts.tv_nsec += 50000000; if(ts.tv_nsec >= 1000000000) { ts.tv_sec++; ts.tv_nsec -= 1000000000; } int res = sem_timedwait(&s, &ts); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } sem_destroy(&s); } else { char __t[512]; snprintf(__t, sizeof(__t), "1");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

