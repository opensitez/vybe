// vybe-test: c/c_while_loop_assignments/while_assignment_array_element
// origin: languages/c/tests/c/test_c_while_loop_assignments.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"10"};
int __n = 1, __i = 0;
 int arr[1] = {2}; while(arr[0]--) { char __t[512]; snprintf(__t, sizeof(__t), "%d", arr[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

