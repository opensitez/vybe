// vybe-test: c/unions/union_character_member_can_feed_putchar_style_format
// origin: languages/c/tests/c/test_unions.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
union Data { int i; char c; };
int main() {
const char *__w[] = {"D\n"};
int __n = 1, __i = 0;
union Data data; data.c = 'D'; { char __t[512]; snprintf(__t, sizeof(__t), "%c\n", data.c);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

