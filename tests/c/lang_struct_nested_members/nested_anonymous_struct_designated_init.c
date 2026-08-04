// vybe-test: c/lang_struct_nested_members/nested_anonymous_struct_designated_init
// origin: languages/c/tests/c/test_lang_struct_nested_members.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Box { struct { int w; int h; } size; };
int main() {
const char *__w[] = {"4 9\n"};
int __n = 1, __i = 0;
struct Box b = {.size = {.w = 4, .h = 9}}; { char __t[512]; snprintf(__t, sizeof(__t), "%d %d\n", b.size.w, b.size.h);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

