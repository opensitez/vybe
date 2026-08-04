// vybe-test: c/pointers_basic/multiple_dereference_reads_array_value
// origin: languages/c/tests/c/test_pointers_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[2] = {8, 9}; int *p = arr;
int main() {
const char *__w[] = {"8 9\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", *p, *(p + 1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

