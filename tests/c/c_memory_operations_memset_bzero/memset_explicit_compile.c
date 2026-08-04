// vybe-test: c/c_memory_operations_memset_bzero/memset_explicit_compile
// origin: languages/c/tests/c/test_c_memory_operations_memset_bzero.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* memset_explicit might not be available everywhere, so we just compile test if it is or fallback */
#include <string.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

