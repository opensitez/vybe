// vybe-test: c/c_posix_semaphores_unnamed/sem_init_in_shm
// origin: languages/c/tests/c/test_c_posix_semaphores_unnamed.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <sys/mman.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 void *p = mmap(NULL, sizeof(sem_t), PROT_READ|PROT_WRITE, MAP_ANON|MAP_SHARED, -1, 0); if(p != MAP_FAILED) { int r = sem_init((sem_t*)p, 1, 1); if (r == 0) { sem_wait((sem_t*)p); sem_post((sem_t*)p); sem_destroy((sem_t*)p); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } else { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } munmap(p, sizeof(sem_t)); } else { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

