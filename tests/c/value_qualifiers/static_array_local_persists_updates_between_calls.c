// vybe-test: c/value_qualifiers/static_array_local_persists_updates_between_calls
// origin: languages/c/tests/c/test_value_qualifiers.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int next_slot(void) { static int values[2] = {1, 2}; values[0] += 1; return values[0]; }
int main() {
const char *__w[] = {"2\n", "3\n"};
int __n = 2, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", next_slot());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", next_slot());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

