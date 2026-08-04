// vybe-test: c/c_flexible_array_members_malloc/fam_malloc_loop
// origin: languages/c/tests/c/test_c_flexible_array_members_malloc.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int len; int data[]; }; int main() {const char *__w[] = {"3"};
int __n = 1, __i = 0;
 struct S *p = malloc(sizeof(struct S) + 3 * sizeof(int)); int sum=0; for(int i=0; i<3; i++) { p->data[i] = i; sum += p->data[i]; } { char __t[512]; snprintf(__t, sizeof(__t), "%d", sum);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(p); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

