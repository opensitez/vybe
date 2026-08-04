// vybe-test: c/c_stdlib_qsort_bsearch/lsearch_basic
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <search.h>
int cmp(const void *a, const void *b) { return (*(int*)a - *(int*)b); }
int main() {const char *__w[] = {"7 4"};
int __n = 1, __i = 0;
 int arr[10] = {9, 2, 5}; size_t n = 3; int key = 7; int *res = lsearch(&key, arr, &n, sizeof(int), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d", *res, (int)n);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

