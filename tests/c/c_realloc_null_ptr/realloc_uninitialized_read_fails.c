// vybe-test: c/c_realloc_null_ptr/realloc_uninitialized_read_fails
// origin: languages/c/tests/c/test_c_realloc_null_ptr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
/* #include <stdlib.h>
int main() { int *p = malloc(sizeof(int)); int *q = realloc(p, 2*sizeof(int)); int val = q[1]; return 0; } */ int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

