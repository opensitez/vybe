// vybe-test: c/nested_functions/forward_declaration_before_definition
// origin: languages/c/tests/c/test_nested_functions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int bar(int x);
int foo(int x) { return bar(x + 1); }
int bar(int x) { return x * 2; }
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", foo(3));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

