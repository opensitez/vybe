// vybe-test: c/c_vla_multidimensional_dynamic/vla_multi_in_loop
// origin: languages/c/tests/c/test_c_vla_multidimensional_dynamic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"5"};
int __n = 1, __i = 0;
 int sum=0; for(int i=1; i<=2; i++) { int arr[i][i]; sum += sizeof(arr)/sizeof(int); } { char __t[512]; snprintf(__t, sizeof(__t), "%d", sum);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

