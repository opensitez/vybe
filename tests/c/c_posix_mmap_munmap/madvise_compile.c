// vybe-test: c/c_posix_mmap_munmap/madvise_compile
// origin: languages/c/tests/c/test_c_posix_mmap_munmap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _BSD_SOURCE
#define _DEFAULT_SOURCE
#include <sys/mman.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 void *p = mmap(NULL, 4096, PROT_READ|PROT_WRITE, MAP_PRIVATE|MAP_ANON, -1, 0); int r = madvise(p, 4096, MADV_NORMAL); { char __t[512]; snprintf(__t, sizeof(__t), "%d", r == 0 || r != 0);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } munmap(p, 4096); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

