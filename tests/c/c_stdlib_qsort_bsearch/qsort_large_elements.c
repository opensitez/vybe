// vybe-test: c/c_stdlib_qsort_bsearch/qsort_large_elements
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
struct L { int a[100]; int k; }; int cmp(const void *a, const void *b) { return ((struct L*)a)->k - ((struct L*)b)->k; }
int main() {const char *__w[] = {"1 2 3"};
int __n = 1, __i = 0;
 struct L arr[3] = {0}; arr[0].k = 3; arr[1].k = 1; arr[2].k = 2; qsort(arr, 3, sizeof(struct L), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", arr[0].k, arr[1].k, arr[2].k);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

