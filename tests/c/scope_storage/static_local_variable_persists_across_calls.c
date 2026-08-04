// vybe-test: c/scope_storage/static_local_variable_persists_across_calls
// origin: languages/c/tests/c/test_scope_storage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int tick(void) { static int n = 0; n++; return n; }
int main() {
const char *__w[] = {"1\n", "2\n"};
int __n = 2, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", tick());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", tick());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

