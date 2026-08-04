// vybe-test: c/c_posix_mmap_munmap/mmap_write_shared
// origin: languages/c/tests/c/test_c_posix_mmap_munmap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"aXcd"};
int __n = 1, __i = 0;
 int fd = open("test_mmap_sh.txt", O_CREAT|O_RDWR, 0644); write(fd, "abcd", 4); void *p = mmap(NULL, 4, PROT_READ|PROT_WRITE, MAP_SHARED, fd, 0); ((char*)p)[1] = 'X'; msync(p, 4, MS_SYNC); munmap(p, 4); lseek(fd, 0, SEEK_SET); char b[5]={0}; read(fd, b, 4); { char __t[512]; snprintf(__t, sizeof(__t), "%s", b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } close(fd); unlink("test_mmap_sh.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

