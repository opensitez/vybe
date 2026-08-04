// vybe-test: c/lang_union_type_punning/struct_with_union_then_read_int_after_char_write
// origin: languages/c/tests/c/test_lang_union_type_punning.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct S { union { int i; char c; } u; };
int main() {
const char *__w[] = {"65\n"};
int __n = 1, __i = 0;
struct S s; s.u.c = 'A'; { char __t[512]; snprintf(__t, sizeof(__t), "%d\n", (int)s.u.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

