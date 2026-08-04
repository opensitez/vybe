// vybe-test: c/c_stdlib_qsort_bsearch/bsearch_empty_array
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int cmp(const void *a, const void *b) { return 0; }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int arr[] = {1}; int key = 1; int *res = bsearch(&key, arr, 0, sizeof(int), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

