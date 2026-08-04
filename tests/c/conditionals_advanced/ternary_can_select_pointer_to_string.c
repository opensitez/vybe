// vybe-test: c/conditionals_advanced/ternary_can_select_pointer_to_string
// origin: languages/c/tests/c/test_conditionals_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"no\n"};
int __n = 1, __i = 0;
int x = 0; char *text = x ? "yes" : "no";
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", text);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

