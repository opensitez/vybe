// vybe-test: c/lang_struct_nested_members/four_level_nesting_access
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct L4 { int v; }; struct L3 { struct L4 l4; }; struct L2 { struct L3 l3; }; struct L1 { struct L2 l2; };
int main() {
const char *__w[] = {"42\n"};
int __n = 1, __i = 0;
struct L1 o = {{{{42}}}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", o.l2.l3.l4.v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

