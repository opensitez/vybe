// vybe-test: c/c_flexible_array_members_malloc/fam_offsetof
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
#include <stddef.h>
struct S { char c; int data[]; }; int main() {const char *__w[] = {"7"};
int __n = 1, __i = 0;
 struct S *p = malloc(sizeof(struct S) + sizeof(int)); p->data[0] = 7; int *d = (int*)((char*)p + offsetof(struct S, data)); { char __t[512]; snprintf(__t, sizeof(__t), "%d", d[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

