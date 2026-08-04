// vybe-test: c/c_flexible_array_members_malloc/fam_malloc_zero_length
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int len; int data[]; }; int main() {const char *__w[] = {"0"};
int __n = 1, __i = 0;
 struct S *p = malloc(sizeof(struct S)); p->len = 0; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p->len);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

