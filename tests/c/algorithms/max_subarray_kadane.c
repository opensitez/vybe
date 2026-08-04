// vybe-test: c/algorithms/max_subarray_kadane
// origin: languages/c/tests/c/test_algorithms.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

int max_subarray(int *a, int n) {
    int max_so_far = a[0], max_ending = a[0];
    for (int i = 1; i < n; i++) {
        max_ending = max_ending + a[i];
        if (max_ending < a[i]) max_ending = a[i];
        if (max_so_far < max_ending) max_so_far = max_ending;
    }
    return max_so_far;
}
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
int a[] = {-2, 1, -3, 4, -1, 2, 1, -5, 4};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", max_subarray(a, 9));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

