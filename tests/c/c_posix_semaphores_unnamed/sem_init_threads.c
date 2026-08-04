// vybe-test: c/c_posix_semaphores_unnamed/sem_init_threads
// origin: languages/c/tests/c/test_c_posix_semaphores_unnamed.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <pthread.h>
#include <unistd.h>
sem_t s;
void* f(void* a) { sem_post(&s); return NULL; }
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int r = sem_init(&s, 0, 0); if(r == 0) { pthread_t t; pthread_create(&t, NULL, f, NULL); sem_wait(&s); pthread_join(t, NULL); sem_destroy(&s); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } else { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

