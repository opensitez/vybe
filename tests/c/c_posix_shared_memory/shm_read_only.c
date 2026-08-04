// vybe-test: c/c_posix_shared_memory/shm_read_only
// origin: languages/c/tests/c/test_c_posix_shared_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int fd1 = shm_open("/test_shm6", O_CREAT | O_RDWR, 0644); ftruncate(fd1, 4096); close(fd1); int fd2 = shm_open("/test_shm6", O_RDONLY, 0644); void *p = mmap(NULL, 4096, PROT_READ | PROT_WRITE, MAP_SHARED, fd2, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p == MAP_FAILED);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd2); shm_unlink("/test_shm6"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

