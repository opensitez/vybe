// vybe-test: c/c_calloc_overflow/calloc_assignment
// origin: languages/c/tests/c/test_c_calloc_overflow.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int main() {const char *__w[] = {"42"};
int __n = 1, __i = 0;
 int *p = calloc(1, sizeof(int)); *p = 42; { char __t[512]; snprintf(__t, sizeof(__t), "%d", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

