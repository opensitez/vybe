// vybe-test: c/structs_advanced/nested_struct_pointer_arrow_can_follow_chain
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Point { int x; int y; }; struct Box { struct Point origin; int size; };
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
struct Box box = {{2, 3}, 4}; struct Box *p = &box;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", p->origin.y);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

