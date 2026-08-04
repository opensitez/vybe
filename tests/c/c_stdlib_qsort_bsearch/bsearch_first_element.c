// vybe-test: c/c_stdlib_qsort_bsearch/bsearch_first_element
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
int cmp(const void *a, const void *b) { return (*(int*)a - *(int*)b); }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int arr[] = {10, 20, 30}; int key = 10; int *res = bsearch(&key, arr, 3, sizeof(int), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res == &arr[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

