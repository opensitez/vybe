// vybe-test: c/algorithms/binary_search_found
// origin: languages/c/tests/c/test_algorithms.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

int bsearch_idx(int *arr, int n, int target) {
    int lo=0, hi=n-1;
    while (lo <= hi) {
        int mid = (lo+hi)/2;
        if (arr[mid] == target) return mid;
        if (arr[mid] < target) lo = mid+1;
        else hi = mid-1;
    }
    return -1;
}
int main() {
const char *__w[] = {"3 -1\n"};
int __n = 1, __i = 0;
int a[] = {1,3,5,7,9,11};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", bsearch_idx(a,6,7), bsearch_idx(a,6,4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

