// vybe-test: c/c_ternary_gnu_omitted_middle/ternary_gnu_omitted_void_fails
// origin: languages/c/tests/c/test_c_ternary_gnu_omitted_middle.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
void f() {} int main() {const char *__w[] = {"ok"};
int __n = 1, __i = 0;
 /* f() ?: 1; // void cannot be tested */ { char __t[512]; snprintf(__t, sizeof(__t), "ok");
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

