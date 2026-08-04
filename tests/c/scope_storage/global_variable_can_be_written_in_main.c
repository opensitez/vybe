// vybe-test: c/scope_storage/global_variable_can_be_written_in_main
// origin: languages/c/tests/c/test_scope_storage.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
int global_value = 7;
int main() {
const char *__w[] = {"9\n"};
int __n = 1, __i = 0;
global_value = 9;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", global_value);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

