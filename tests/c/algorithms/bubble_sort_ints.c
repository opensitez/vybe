// vybe-test: c/algorithms/bubble_sort_ints
// origin: languages/c/tests/c/test_algorithms.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

void bubble_sort(int *arr, int n) {
    for (int i = 0; i < n-1; i++)
        for (int j = 0; j < n-1-i; j++)
            if (arr[j] > arr[j+1]) {
                int t = arr[j]; arr[j] = arr[j+1]; arr[j+1] = t;
            }
}
int main() {
const char *__w[] = {"1 2 3 5 8 9\n"};
int __n = 1, __i = 0;
int a[] = {5,3,8,1,9,2};
bubble_sort(a, 6);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d %d %d %d\n", a[0],a[1],a[2],a[3],a[4],a[5]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

