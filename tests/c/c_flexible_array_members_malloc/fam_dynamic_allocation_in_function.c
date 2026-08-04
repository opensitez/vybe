// vybe-test: c/c_flexible_array_members_malloc/fam_dynamic_allocation_in_function
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int len; int data[]; }; struct S *create(int n) { struct S *p = malloc(sizeof(struct S) + n * sizeof(int)); p->len = n; return p; } int main() {const char *__w[] = {"4"};
int __n = 1, __i = 0;
 struct S *p = create(2); p->data[1] = 4; { char __t[512]; snprintf(__t, sizeof(__t), "%d", p->data[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

