// vybe-test: c/c_stdlib_qsort_bsearch/qsort_strings
// origin: languages/c/tests/c/test_c_stdlib_qsort_bsearch.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <stdlib.h>
#include <string.h>
int cmp(const void *a, const void *b) { return strcmp(*(const char**)a, *(const char**)b); }
int main() {const char *__w[] = {"apple mango zebra"};
int __n = 1, __i = 0;
 const char *arr[] = {"zebra", "apple", "mango"}; qsort(arr, 3, sizeof(char*), cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%s %s %s", arr[0], arr[1], arr[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

