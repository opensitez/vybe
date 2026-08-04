// vybe-test: c/pointers_basic/pointer_indirection_on_array_slot_can_update_element
// origin: languages/c/tests/c/test_pointers_basic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int arr[2] = {1, 2}; int *p = &arr[1];
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
*p = 9;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

