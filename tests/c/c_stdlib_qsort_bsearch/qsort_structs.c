// vybe-test: c/c_stdlib_qsort_bsearch/qsort_structs
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct S { int k; char v; }; int cmp(const void *a, const void *b) { return ((struct S*)a)->k - ((struct S*)b)->k; }
int main() {const char *__w[] = {"a c e"};
int __n = 1, __i = 0;
 struct S arr[] = {{5,'e'}, {1,'a'}, {3,'c'}}; qsort(arr, 3, sizeof(struct S), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%c %c %c", arr[0].v, arr[1].v, arr[2].v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

