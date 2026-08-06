// vybe-test: c/includes/include_pragma_once_twice
#include <stdio.h>
#include <string.h>
#include <assert.h>
#include "inc_once.h"
#include "inc_once.h"
int main() {
const char *__w[] = {"3"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d", ONCE_TOKEN);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}
