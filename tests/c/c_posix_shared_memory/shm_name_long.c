// vybe-test: c/c_posix_shared_memory/shm_name_long
// origin: languages/c/tests/c/test_c_posix_shared_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 char name[300]; name[0] = '/'; for(int i=1; i<250; i++) name[i] = 'a'; name[250] = 0; int fd = shm_open(name, O_CREAT | O_RDWR, 0644); { char __t[512]; snprintf(__t, sizeof(__t), "%d", fd == -1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

