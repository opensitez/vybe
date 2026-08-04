// vybe-test: c/structs_advanced/struct_with_pointer_field_can_follow_suffix
// origin: languages/c/tests/c/test_structs_advanced.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
struct Slice { char *text; };
int main() {
const char *__w[] = {"llo\n"};
int __n = 1, __i = 0;
struct Slice slice = {"hello" + 2};
{ char __t[512]; snprintf(__t, sizeof(__t), "%s\n", slice.text);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

