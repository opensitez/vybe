// vybe-test: c/structs_advanced/struct_can_be_passed_to_function_by_value
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; }; int sum_pair(struct Pair pair) { return pair.a + pair.b; }
int main() {
const char *__w[] = {"11\n"};
int __n = 1, __i = 0;
struct Pair pair = {5, 6};
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", sum_pair(pair));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

