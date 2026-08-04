// vybe-test: c/pointers_basic/pointer_can_point_to_array_first_element
// origin: languages/c/tests/c/test_pointers_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[3] = {4, 5, 6}; int *p = arr;
int main() {
const char *__w[] = {"4\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

