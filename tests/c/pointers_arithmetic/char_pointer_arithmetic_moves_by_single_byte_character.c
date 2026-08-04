// vybe-test: c/pointers_arithmetic/char_pointer_arithmetic_moves_by_single_byte_character
// origin: languages/c/tests/c/test_pointers_arithmetic.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>
char text[] = "hello"; char *p = text;
int main() {
const char *__w[] = {"l\n"};
int __n = 1, __i = 0;
p += 2;
{ char __t[512]; snprintf(__t, sizeof(__t), "%c\n", *p);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

