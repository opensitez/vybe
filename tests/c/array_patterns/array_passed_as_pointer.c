// vybe-test: c/array_patterns/array_passed_as_pointer
// origin: languages/c/tests/c/test_array_patterns.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int first(int *a) { return a[0]; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
int arr[3] = {10,20,30};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", first(arr));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

