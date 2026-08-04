// vybe-test: c/lang_union_type_punning/union_in_struct_write_int_field
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
union U { int i; char c; }; struct Box { union U data; };
int main() {
const char *__w[] = {"60\n"};
int __n = 1, __i = 0;
struct Box b; b.data.i = 60; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", b.data.i);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

