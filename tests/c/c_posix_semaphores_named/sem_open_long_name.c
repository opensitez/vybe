// vybe-test: c/c_posix_semaphores_named/sem_open_long_name
// origin: languages/c/tests/c/test_c_posix_semaphores_named.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 char name[300]; name[0] = '/'; for(int i=1; i<250; i++) name[i] = 'a'; name[250] = 0; sem_t *s = sem_open(name, O_CREAT, 0644, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%d", s == SEM_FAILED);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } /* Most implementations reject very long names */ if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

