// vybe-test: c/unions/union_pointer_can_write_character_member
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; };
int main() {
const char *__w[] = {"B\n"};
int __n = 1, __i = 0;
union Data data; union Data *p = &data; p->c = 'B'; { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", data.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

