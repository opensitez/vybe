// vybe-test: c/c_stdlib_qsort_bsearch/qsort_r_gnu
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <stdlib.h>
int cmp(const void *a, const void *b, void *arg) { int dir = *(int*)arg; return (*(int*)a - *(int*)b) * dir; }
int main() {const char *__w[] = {"3 2 1"};
int __n = 1, __i = 0;
 int arr[] = {3, 1, 2}; int dir = -1; qsort_r(arr, 3, sizeof(int), cmp, &dir); { char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d", arr[0], arr[1], arr[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

