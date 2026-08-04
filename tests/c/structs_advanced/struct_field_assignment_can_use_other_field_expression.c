// vybe-test: c/structs_advanced/struct_field_assignment_can_use_other_field_expression
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Pair { int a; int b; };
int main() {
const char *__w[] = {"7\n"};
int __n = 1, __i = 0;
struct Pair pair = {2, 0};
pair.b = pair.a + 5;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", pair.b);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

