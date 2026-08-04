// vybe-test: c/c_aligned_alloc_usage/aligned_alloc_zero_size
// origin: languages/c/tests/c/test_c_aligned_alloc_usage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 void *p = aligned_alloc(16, 0); if (p) free(p); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

