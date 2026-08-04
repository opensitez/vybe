// vybe-test: c/structs_advanced/struct_copy_is_independent_after_mutation
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; };
int main() {
const char *__w[] = {"1 9\n"};
int __n = 1, __i = 0;
struct Pair first = {1, 2}; struct Pair second = first; second.a = 9;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", first.a, second.a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

