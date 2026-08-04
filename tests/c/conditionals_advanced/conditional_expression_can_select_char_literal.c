// vybe-test: c/conditionals_advanced/conditional_expression_can_select_char_literal
// origin: languages/c/tests/c/test_conditionals_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int main() {
const char *__w[] = {"y\n"};
int __n = 1, __i = 0;
int x = 1;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", x ? 'y' : 'n');
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

