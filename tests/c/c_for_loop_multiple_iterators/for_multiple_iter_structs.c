// vybe-test: c/c_for_loop_multiple_iterators/for_multiple_iter_structs
// origin: languages/c/tests/c/test_c_for_loop_multiple_iterators.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct S { int x; }; int main() {const char *__w[] = {"2"};
int __n = 1, __i = 0;
 for(struct S s1={1}, s2={2}; s1.x<2; s1.x++) { char __t[512]; snprintf(__t, sizeof(__t), "%d", s2.x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0; }

