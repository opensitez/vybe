// vybe-test: c/const_correctness/const_function_parameter
// origin: languages/c/tests/c/test_const_correctness.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int strlen_safe(const char *s) { int n = 0; while (*s++) n++; return n; }
int main() {
const char *__w[] = {"5\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", strlen_safe("hello"));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

