// vybe-test: c/c_logical_short_circuit_side_effects/short_circuit_array_bounds
// origin: languages/c/tests/c/test_c_logical_short_circuit_side_effects.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"no"};
int __n = 1, __i = 0;
 int arr[2]={1,2}; int i=5; if (i<2 && arr[i]==1) { char __t[512]; snprintf(__t, sizeof(__t), "yes");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } else { char __t[512]; snprintf(__t, sizeof(__t), "no");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

