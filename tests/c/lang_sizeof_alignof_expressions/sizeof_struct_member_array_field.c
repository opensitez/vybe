// vybe-test: c/lang_sizeof_alignof_expressions/sizeof_struct_member_array_field
// origin: languages/c/tests/c/test_lang_sizeof_alignof_expressions.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Buf { char data[8]; };
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
struct Buf b; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)sizeof(b.data));
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

