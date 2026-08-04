// vybe-test: c/unions/union_can_live_inside_struct_field
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; }; struct Box { union Data data; };
int main() {
const char *__w[] = {"90\n"};
int __n = 1, __i = 0;
struct Box box; box.data.i = 90; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", box.data.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

