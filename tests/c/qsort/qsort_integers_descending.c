// vybe-test: c/qsort/qsort_integers_descending
// origin: languages/c/tests/c/test_qsort.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int cmp_desc(const void *a, const void *b) { return *(int*)b - *(int*)a; }
int main() {
const char *__w[] = {"5 4 3 1 1\n"};
int __n = 1, __i = 0;

int arr[] = {3, 1, 4, 1, 5};
qsort(arr, 5, sizeof(int), cmp_desc);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d %d\n", arr[0], arr[1], arr[2], arr[3], arr[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

