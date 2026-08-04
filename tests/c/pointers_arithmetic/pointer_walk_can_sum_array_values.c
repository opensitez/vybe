// vybe-test: c/pointers_arithmetic/pointer_walk_can_sum_array_values
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[4] = {1, 2, 3, 4}; int *p = arr;
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int total = 0;
for (int i = 0; i < 4; i++) total += *(p + i);
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", total);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

