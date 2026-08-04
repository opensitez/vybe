// vybe-test: c/c_posix_shared_memory/shm_shared_between_processes
// origin: languages/c/tests/c/test_c_posix_shared_memory.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
#include <sys/wait.h>
int main() {const char *__w[] = {"P"};
int __n = 1, __i = 0;
 int fd = shm_open("/test_shm8", O_CREAT | O_RDWR, 0644); ftruncate(fd, 4096); void *p = mmap(NULL, 4096, PROT_READ | PROT_WRITE, MAP_SHARED, fd, 0); pid_t pid = fork(); if(pid==0) { ((char*)p)[0] = 'P'; msync(p, 4096, MS_SYNC); _exit(0); } wait(NULL); { char __t[512]; snprintf(__t, sizeof(__t), "%c", ((char*)p)[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } munmap(p, 4096); close(fd); shm_unlink("/test_shm8"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

