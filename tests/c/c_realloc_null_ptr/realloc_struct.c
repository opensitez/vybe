// vybe-test: c/c_realloc_null_ptr/realloc_struct
// origin: languages/c/tests/c/test_c_realloc_null_ptr.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int a; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct S *p = malloc(sizeof(struct S)); p->a = 1; p = realloc(p, 2 * sizeof(struct S)); p[1].a = 2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p[0].a + p[1].a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

