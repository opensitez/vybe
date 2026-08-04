// vybe-test: c/lang_struct_union_enum/typedef_struct_tag
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
typedef struct { int v; } Box;
int main() {
const char *__w[] = {"3\n"};
int __n = 1, __i = 0;
Box b = {3}; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", b.v);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

