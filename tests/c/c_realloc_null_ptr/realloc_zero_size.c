// vybe-test: c/c_realloc_null_ptr/realloc_zero_size
// origin: languages/c/tests/c/test_c_realloc_null_ptr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 int *p = malloc(sizeof(int)); void *q = realloc(p, 0); { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } /* standard C says behavior is implementation-defined, might free or return NULL or unique ptr */ if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

