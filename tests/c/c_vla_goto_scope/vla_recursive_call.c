// vybe-test: c/c_vla_goto_scope/vla_recursive_call
// origin: languages/c/tests/c/test_c_vla_goto_scope.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int f(int n) { if(n==0) return 0; int arr[n]; arr[n-1]=n; return arr[n-1] + f(n-1); } int main() {const char *__w[] = {"6"};
int __n = 1, __i = 0;
 { char __t[512]; snprintf(__t, sizeof(__t), "%d", f(3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

