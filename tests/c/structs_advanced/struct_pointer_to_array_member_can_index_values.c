// vybe-test: c/structs_advanced/struct_pointer_to_array_member_can_index_values
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Row { int values[3]; };
int main() {
const char *__w[] = {"6\n"};
int __n = 1, __i = 0;
struct Row row = {{2, 4, 6}}; struct Row *p = &row;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->values[2]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

