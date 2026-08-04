// vybe-test: c/c_realloc_null_ptr/realloc_fail
// origin: languages/c/tests/c/test_c_realloc_null_ptr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
#include <stdint.h>
int main() {const char *__w[] = {"fail"};
int __n = 1, __i = 0;
 void *p = malloc(10); void *q = realloc(p, SIZE_MAX); if (q == NULL) { char __t[512]; snprintf(__t, sizeof(__t), "fail");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else free(q); free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

