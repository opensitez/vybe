// vybe-test: c/functions_advanced/function_can_return_early_from_branch
// origin: languages/c/tests/c/test_functions_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int sign_label(int x) { if (x < 0) return -1; if (x > 0) return 1; return 0; }
int main() {
const char *__w[] = {"-1 0 1\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d %d\n", sign_label(-4), sign_label(0), sign_label(4));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

