// vybe-test: c/c_stdlib_qsort_bsearch/tsearch_tfind_basic
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#define _GNU_SOURCE
#include <search.h>
#include <stdlib.h>
int cmp(const void *a, const void *b) { return (*(int*)a - *(int*)b); }
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 void *root = NULL; int x=3, y=1, z=2; tsearch(&x, &root, cmp); tsearch(&y, &root, cmp); tsearch(&z, &root, cmp); int key=2; void *res = tfind(&key, &root, cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d", res != NULL);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } tdestroy(root, free); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

