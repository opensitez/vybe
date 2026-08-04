// vybe-test: c/c_malloc_zero_bytes/malloc_struct_with_pointer
// origin: languages/c/tests/c/test_c_malloc_zero_bytes.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int *p; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct S *s = malloc(sizeof(struct S)); s->p = malloc(sizeof(int)); *s->p = 3; { char __t[512]; snprintf(__t, sizeof(__t), "%d", *s->p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(s->p); free(s); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

