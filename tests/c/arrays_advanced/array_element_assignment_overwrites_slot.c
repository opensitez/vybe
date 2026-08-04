// vybe-test: c/arrays_advanced/array_element_assignment_overwrites_slot
// origin: languages/c/tests/c/test_arrays_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[3] = {1, 2, 3};
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
arr[1] = 9;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

