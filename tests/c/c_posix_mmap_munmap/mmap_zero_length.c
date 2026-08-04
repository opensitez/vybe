// vybe-test: c/c_posix_mmap_munmap/mmap_zero_length
// origin: languages/c/tests/c/test_c_posix_mmap_munmap.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _POSIX_C_SOURCE 200809L
#include <sys/mman.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 void *p = mmap(NULL, 0, PROT_READ, MAP_PRIVATE|MAP_ANON, -1, 0); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p == MAP_FAILED);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

