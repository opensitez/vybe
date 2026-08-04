// vybe-test: c/structs_advanced/struct_pointer_can_iterate_array_of_structs
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; };
int main() {
const char *__w[] = {"2 3\n"};
int __n = 1, __i = 0;
struct Pair pairs[2] = {{1, 2}, {3, 4}}; struct Pair *p = pairs;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", p[0].b, p[1].a);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

