// vybe-test: c/c_posix_semaphores_named/sem_open_limit
// origin: languages/c/tests/c/test_c_posix_semaphores_named.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <semaphore.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 /* Check we can open at least 5 */ sem_t *s[5]; int ok = 1; char name[20]; for(int i=0; i<5; i++) { sprintf(name, "/test_sem_l%d", i); s[i] = sem_open(name, O_CREAT, 0644, 1); if(s[i] == SEM_FAILED) ok = 0; } { char __t[512]; snprintf(__t, sizeof(__t), "%d", ok);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } for(int i=0; i<5; i++) { sprintf(name, "/test_sem_l%d", i); if(s[i] != SEM_FAILED) { sem_close(s[i]); sem_unlink(name); } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

