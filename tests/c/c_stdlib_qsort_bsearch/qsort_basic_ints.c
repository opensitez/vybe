// vybe-test: c/c_stdlib_qsort_bsearch/qsort_basic_ints
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int cmp(const void *a, const void *b) { return (*(int*)a - *(int*)b); }
int main() {const char *__w[] = {"1 1 9"};
int __n = 1, __i = 0;
 int arr[] = {3, 1, 4, 1, 5, 9}; qsort(arr, 6, sizeof(int), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", arr[0], arr[1], arr[5]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

