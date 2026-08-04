// vybe-test: c/c_stdlib_memory_alignment/mallopt_check
// origin: languages/c/tests/c/test_c_stdlib_memory_alignment.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <malloc.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int res = mallopt(M_TRIM_THRESHOLD, 128*1024); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == 1);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

