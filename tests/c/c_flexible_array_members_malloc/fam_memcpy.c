// vybe-test: c/c_flexible_array_members_malloc/fam_memcpy
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
#include <string.h>
struct S { int len; int data[]; }; int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 struct S *p1 = malloc(sizeof(struct S) + 2 * sizeof(int)); p1->len=2; p1->data[0]=1; p1->data[1]=2; struct S *p2 = malloc(sizeof(struct S) + 2 * sizeof(int)); memcpy(p2, p1, sizeof(struct S) + 2 * sizeof(int)); { char __t[512]; snprintf(__t, sizeof(__t), "%d", p2->data[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p1); free(p2); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

