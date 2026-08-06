// vybe-test: c/includes/include_nested_chain
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include "inc_outer.h"
int main() {
const char *__w[] = {"43 42"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d", OUTER_SUM, BASICS_MAGIC);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
