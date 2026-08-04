// vybe-test: c/c_posix_shared_memory/shm_mmap_executable
// origin: languages/c/tests/c/test_c_posix_shared_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int fd = shm_open("/test_shm15", O_CREAT | O_RDWR, 0644); ftruncate(fd, 4096); void *p = mmap(NULL, 4096, PROT_READ | PROT_EXEC, MAP_SHARED, fd, 0); /* May fail or succeed depending on OS/mount options, check compile and basic run */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(p != MAP_FAILED) munmap(p, 4096); close(fd); shm_unlink("/test_shm15"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

