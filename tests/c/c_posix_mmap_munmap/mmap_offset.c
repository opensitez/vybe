// vybe-test: c/c_posix_mmap_munmap/mmap_offset
// origin: languages/c/tests/c/test_c_posix_mmap_munmap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
#include <fcntl.h>
#include <unistd.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 long ps = sysconf(_SC_PAGESIZE); int fd = open("test_off.txt", O_CREAT|O_RDWR, 0644); lseek(fd, ps + 4, SEEK_SET); write(fd, "x", 1); void *p = mmap(NULL, 4, PROT_READ, MAP_PRIVATE, fd, ps); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p != MAP_FAILED);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if(p != MAP_FAILED) munmap(p, 4); close(fd); unlink("test_off.txt"); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

