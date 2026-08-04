// vybe-test: c/lang_struct_nested_members/anonymous_struct_in_struct_direct_access
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Pair { struct { int lo; int hi; } range; };
int main() {
const char *__w[] = {"3 8\n"};
int __n = 1, __i = 0;
struct Pair p; p.range.lo = 3; p.range.hi = 8; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", p.range.lo, p.range.hi);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

