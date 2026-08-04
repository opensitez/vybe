// vybe-test: c/typedefs/typedef_alias_for_nested_struct_can_copy_value
// origin: languages/c/tests/c/test_typedefs.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
typedef struct { int a; int b; } Pair;
int main() {
const char *__w[] = {"1 2\n"};
int __n = 1, __i = 0;
Pair first = {1, 2}; Pair second = first; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", second.a, second.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

