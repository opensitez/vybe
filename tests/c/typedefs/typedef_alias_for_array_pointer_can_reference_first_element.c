// vybe-test: c/typedefs/typedef_alias_for_array_pointer_can_reference_first_element
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int *IntPtr;
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
int arr[2] = {5, 6}; IntPtr ptr = &arr[0]; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *ptr);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

