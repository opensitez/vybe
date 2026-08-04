// vybe-test: c/c_posix_semaphores_named/sem_open_exclusive
// origin: languages/c/tests/c/test_c_posix_semaphores_named.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 sem_t *s1 = sem_open("/test_sem8", O_CREAT, 0644, 1); sem_t *s2 = sem_open("/test_sem8", O_CREAT | O_EXCL, 0644, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", s2 == SEM_FAILED);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } sem_close(s1); sem_unlink("/test_sem8"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

