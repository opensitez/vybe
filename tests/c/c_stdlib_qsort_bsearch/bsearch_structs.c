// vybe-test: c/c_stdlib_qsort_bsearch/bsearch_structs
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int id; char n; }; int cmp(const void *a, const void *b) { return ((struct S*)a)->id - ((struct S*)b)->id; }
int main() {const char *__w[] = {"B"};
int __n = 1, __i = 0;
 struct S arr[] = {{1,'A'}, {2,'B'}, {3,'C'}}; struct S key = {2, 0}; struct S *res = bsearch(&key, arr, 3, sizeof(struct S), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%c", res->n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

