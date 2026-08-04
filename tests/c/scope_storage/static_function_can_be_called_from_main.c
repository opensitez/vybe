// vybe-test: c/scope_storage/static_function_can_be_called_from_main
// origin: languages/c/tests/c/test_scope_storage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
static int twice(int x) { return x * 2; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", twice(5));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

