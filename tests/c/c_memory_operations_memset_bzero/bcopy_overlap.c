// vybe-test: c/c_memory_operations_memset_bzero/bcopy_overlap
// origin: languages/c/tests/c/test_c_memory_operations_memset_bzero.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <strings.h>
int main() {const char *__w[] = {"ababcd"};
int __n = 1, __i = 0;
 char s[] = "abcdef"; bcopy(s, s + 2, 4); { char __t[512]; snprintf(__t, sizeof(__t), "%s", s);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

