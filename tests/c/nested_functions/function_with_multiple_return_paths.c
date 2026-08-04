// vybe-test: c/nested_functions/function_with_multiple_return_paths
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
const char* classify(int n) { if (n < 0) return "neg"; if (n == 0) return "zero"; return "pos"; }
int main() {
const char *__w[] = {"neg zero pos\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%s %s %s\n", classify(-1), classify(0), classify(1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

