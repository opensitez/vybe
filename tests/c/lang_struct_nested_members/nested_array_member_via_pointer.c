// vybe-test: c/lang_struct_nested_members/nested_array_member_via_pointer
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Pt { int x; }; struct Row { struct Pt cells[3]; };
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
struct Row r = {{{1},{2},{3}}}; struct Row *rp = &r; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", rp->cells[2].x);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

