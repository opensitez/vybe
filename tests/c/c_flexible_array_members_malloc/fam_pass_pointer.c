// vybe-test: c/c_flexible_array_members_malloc/fam_pass_pointer
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static const char *__w[] = {"88"};
static int __n = 1, __i = 0;
#include <stdlib.h>
struct S { int len; int data[]; }; void f(struct S *p) { { char __t[512]; snprintf(__t, sizeof(__t), "%d", p->data[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } int main() { struct S *p = malloc(sizeof(struct S) + sizeof(int)); p->data[0] = 88; f(p); free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

