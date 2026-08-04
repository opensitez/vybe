// vybe-test: c/cover_stdlib_h/qsort_sorts
// origin: languages/c/tests/c/test_cover_stdlib_h.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
#include <stdlib.h>
int cmp(const void *a, const void *b) { return *(int*)a - *(int*)b; }
int main() {
const char *__w[] = {"2\n"};
int __n = 1, __i = 0;
int a[]={3,1,2}; qsort(a,3,sizeof(int),cmp); { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", a[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

