// vybe-test: c/algorithms/insertion_sort_ints
// origin: languages/c/tests/c/test_algorithms.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

void insertion_sort(int *arr, int n) {
    for (int i = 1; i < n; i++) {
        int key = arr[i], j = i - 1;
        while (j >= 0 && arr[j] > key) { arr[j+1] = arr[j]; j--; }
        arr[j+1] = key;
    }
}
int main() {
const char *__w[] = {"1 2 4 5 7\n"};
int __n = 1, __i = 0;
int a[] = {4,2,7,1,5};
insertion_sort(a, 5);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d %d\n", a[0],a[1],a[2],a[3],a[4]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

