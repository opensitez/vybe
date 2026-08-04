// vybe-test: c/c_posix_shared_memory/shm_mmap_private
// origin: languages/c/tests/c/test_c_posix_shared_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"x"};
int __n = 1, __i = 0;
 int fd = shm_open("/test_shm12", O_CREAT | O_RDWR, 0644); ftruncate(fd, 4096); write(fd, "x", 1); void *p = mmap(NULL, 4096, PROT_READ | PROT_WRITE, MAP_PRIVATE, fd, 0); ((char*)p)[0] = 'y'; lseek(fd, 0, SEEK_SET); char b[2]={0}; read(fd, b, 1); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } munmap(p, 4096); close(fd); shm_unlink("/test_shm12"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

