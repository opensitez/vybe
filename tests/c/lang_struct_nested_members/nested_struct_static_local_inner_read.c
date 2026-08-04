// vybe-test: c/lang_struct_nested_members/nested_struct_static_local_inner_read
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Inner { int n; }; struct Outer { struct Inner in; }; int read_cached(void) { static struct Outer cache = {{10}}; return cache.in.n; }
int main() {
const char *__w[] = {"10\n"};
int __n = 1, __i = 0;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", read_cached());
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

