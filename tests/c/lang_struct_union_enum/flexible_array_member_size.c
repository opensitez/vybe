// vybe-test: c/lang_struct_union_enum/flexible_array_member_size
// origin: languages/c/tests/c/test_lang_struct_union_enum.rs
#include <string.h>
#include <assert.h>
#include <stdio.h>
struct Buf { int n; char data[]; };
int main() {
const char *__w[] = {"a\n"};
int __n = 1, __i = 0;
struct Buf *b = malloc(sizeof(struct Buf) + 4); b->n = 4; b->data[0]='a'; { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", b->data[0]);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; } free(b); if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

