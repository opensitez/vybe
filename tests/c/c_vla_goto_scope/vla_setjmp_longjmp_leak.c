// vybe-test: c/c_vla_goto_scope/vla_setjmp_longjmp_leak
// origin: languages/c/tests/c/test_c_vla_goto_scope.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include <setjmp.h>
int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 jmp_buf env; if(setjmp(env) == 0) { int n=1000; int arr[n]; longjmp(env, 1); } { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

