// vybe-test: c/c_posix_mmap_munmap/mlock_munlock
// origin: languages/c/tests/c/test_c_posix_mmap_munmap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 void *p = mmap(NULL, 4096, PROT_READ|PROT_WRITE, MAP_PRIVATE|MAP_ANON, -1, 0); int r1 = mlock(p, 4096); if(r1 == 0) munlock(p, 4096); /* May fail if not root or limits exceeded, we just check it compiles */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } munmap(p, 4096); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

