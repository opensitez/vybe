// vybe-test: c/c_vla_goto_scope/vla_nested_blocks_same_name
// origin: languages/c/tests/c/test_c_vla_goto_scope.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 int n=1; { int arr[n]; arr[0]=1; { int arr[n+1]; arr[1]=2; { char __t[512]; snprintf(__t, sizeof(__t), "%d", arr[1]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } } } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

