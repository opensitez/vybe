// vybe-test: c/scope_storage/local_variable_lifetime_is_per_call
// origin: languages/c/tests/c/test_scope_storage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int next(int x) { int y = x + 1; return y; }
int main() {
const char *__w[] = {"2 2\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", next(1), next(1));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

