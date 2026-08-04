// vybe-test: c/typedefs/typedef_alias_for_pointer_to_function_returning_pointer_can_call
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef int *(*IdFn)(int *); int *identity(int *p) { return p; }
int main() {
const char *__w[] = {"12\n"};
int __n = 1, __i = 0;
int value = 12; IdFn fn = identity; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", *fn(&value));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

