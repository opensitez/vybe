// vybe-test: c/includes/include_guard_double_include
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include "inc_guarded.h"
#include "inc_guarded.h"
int main() {
const char *__w[] = {"7"};
int __n = 1, __i = 0;
GUARD_COUNTER_TYPE v = GUARD_VALUE;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d", v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
