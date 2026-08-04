// vybe-test: c/c_vla_goto_scope/vla_goto_break_scope
// origin: languages/c/tests/c/test_c_vla_goto_scope.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {const char *__w[] = {"1"};
int __n = 1, __i = 0;
 int sum=0; for(int i=1; i<=2; i++) { int arr[i]; arr[i-1]=i; if (i==2) break; sum += arr[i-1]; } { char __t[512]; snprintf(__t, sizeof(__t), "%d", sum);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

