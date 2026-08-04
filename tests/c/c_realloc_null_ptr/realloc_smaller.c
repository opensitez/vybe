// vybe-test: c/c_realloc_null_ptr/realloc_smaller
// origin: languages/c/tests/c/test_c_realloc_null_ptr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"9"};
int __n = 1, __i = 0;
 int *p = malloc(20 * sizeof(int)); for(int i=0; i<20; i++) p[i]=i; int *q = realloc(p, 10 * sizeof(int)); { char __t[512]; snprintf(__t, sizeof(__t), "%d", q[9]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(q); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

